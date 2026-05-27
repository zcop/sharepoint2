<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Service;

use OCA\Files_External\Lib\InsufficientDataForMeaningfulAnswerException;
use OCA\Files_External\Lib\StorageConfig;
use OCA\Files_External\NotFoundException;
use OCA\Files_External\Service\GlobalStoragesService;
use OCP\Files\Storage\IStorage;
use OCP\IConfig;
use OCP\Lock\LockedException;
use OCP\Files\StorageNotAvailableException;
use Psr\Log\LoggerInterface;

class CacheWarmupService {
    public function __construct(
        private GlobalStoragesService $globalStoragesService,
        private CacheWarmupStateService $stateService,
        private IConfig $config,
        private LoggerInterface $logger,
    ) {
    }

    /**
     * @return array{mode:string,mount_id:int,level1_dirs:int,scanned_dirs:int,full_scan:bool,path:string}
     */
    public function warmupMountById(int $mountId, bool $fullScan = false, string $path = ''): array {
        try {
            $storageConfig = $this->globalStoragesService->getStorage($mountId);
        } catch (NotFoundException $e) {
            throw new \RuntimeException('Mount not found: ' . $mountId, 0, $e);
        }

        if (!$storageConfig instanceof StorageConfig) {
            throw new \RuntimeException('Mount not found: ' . $mountId);
        }

        if ($storageConfig->getBackend()->getIdentifier() !== 'sharepoint2') {
            throw new \RuntimeException('Mount ' . $mountId . ' is not sharepoint2');
        }

        $storage = $this->createStorage($storageConfig);
        if ($storage === null) {
            throw new \RuntimeException('Failed to construct storage for mount ' . $mountId);
        }

        return $this->warmupStorage($mountId, $storageConfig, $storage, $fullScan, $path);
    }

    /**
     * @return array{mode:string,mount_id:int,level1_dirs:int,scanned_dirs:int,full_scan:bool,path:string}
     */
    private function warmupStorage(int $mountId, StorageConfig $storageConfig, IStorage $storage, bool $fullScan, string $path): array {
        $mountKey = $this->buildMountKeyFromOptions($storageConfig->getBackendOptions());
        $scanner = $storage->getScanner();

        if ($fullScan) {
            $this->logger->info('Sharepoint2 cache warmup: full scan start', ['mountId' => $mountId]);
            $scanner->scan('');
            $this->stateService->finishWarmup($mountKey);
            return [
                'mode' => 'full',
                'mount_id' => $mountId,
                'level1_dirs' => 0,
                'scanned_dirs' => 0,
                'full_scan' => true,
                'path' => '',
            ];
        }

        $path = trim($path, '/');
        if ($path !== '') {
            $this->logger->info('Sharepoint2 cache warmup: path scan start', [
                'mountId' => $mountId,
                'path' => $path,
            ]);
            $scanner->scan($path);
            $this->stateService->markLevel1Done($mountKey, $path);
            return [
                'mode' => 'path',
                'mount_id' => $mountId,
                'level1_dirs' => 0,
                'scanned_dirs' => 1,
                'full_scan' => false,
                'path' => $path,
            ];
        }

        // Stage 1: scan root
        $this->logger->info('Sharepoint2 cache warmup: root scan start', ['mountId' => $mountId]);
        $scanner->scan('');

        // Stage 2: scan each level-1 directory
        $level1Dirs = $this->listLevel1Directories($storage);
        $this->stateService->startWarmup($mountKey, $mountId, $level1Dirs);

        $scanned = 0;
        try {
            foreach ($level1Dirs as $dir) {
                $scanner->scan($dir);
                $this->stateService->markLevel1Done($mountKey, $dir);
                $scanned++;
            }
            $this->stateService->finishWarmup($mountKey);
        } catch (LockedException $e) {
            $this->stateService->failWarmup($mountKey, 'Scanner lock: ' . $e->getMessage());
            throw $e;
        } catch (\Throwable $e) {
            $this->stateService->failWarmup($mountKey, $e->getMessage());
            throw $e;
        }

        return [
            'mode' => 'warmup',
            'mount_id' => $mountId,
            'level1_dirs' => count($level1Dirs),
            'scanned_dirs' => $scanned,
            'full_scan' => false,
            'path' => '',
        ];
    }

    /**
     * @return list<string>
     */
    private function listLevel1Directories(IStorage $storage): array {
        $result = [];
        $content = $storage->getDirectoryContent('');
        foreach ($content as $item) {
            if (!is_array($item)) {
                continue;
            }
            $name = isset($item['name']) ? trim((string)$item['name']) : '';
            if ($name === '') {
                continue;
            }

            $type = (string)($item['type'] ?? '');
            if ($type === 'dir' || isset($item['folder'])) {
                $result[] = trim($name, '/');
            }
        }

        return array_values(array_unique($result));
    }

    private function createStorage(StorageConfig $storageConfig): ?IStorage {
        try {
            $storageConfig->getAuthMechanism()->manipulateStorageConfig($storageConfig, null);
            $storageConfig->getBackend()->manipulateStorageConfig($storageConfig, null);
        } catch (InsufficientDataForMeaningfulAnswerException|StorageNotAvailableException $e) {
            $this->logger->warning('Sharepoint2 cache warmup: insufficient auth/backend data', [
                'error' => $e->getMessage(),
            ]);
        }

        try {
            $class = $storageConfig->getBackend()->getStorageClass();
            /** @var IStorage $storage */
            $storage = new $class($storageConfig->getBackendOptions());
            if (!$storage->test()) {
                return null;
            }
            return $storage;
        } catch (\Throwable $e) {
            $this->logger->warning('Sharepoint2 cache warmup: failed to instantiate storage', [
                'error' => $e->getMessage(),
            ]);
            return null;
        }
    }

    private function buildMountKeyFromOptions(array $options): string {
        $siteUrl = rtrim((string)($options['site_url'] ?? ''), '/');
        $library = trim((string)($options['library'] ?? ''), '/');
        $clientId = (string)($options['client_id'] ?? '');

        $tenant = trim((string)($options['tenant'] ?? ''));
        if ($tenant === '') {
            $tenant = trim((string)$this->config->getSystemValue('sharepoint2_tenant', ''));
        }
        if ($tenant === '') {
            $tenant = 'common';
        }

        return CacheWarmupStateService::buildMountKey($siteUrl, $library, $tenant, $clientId);
    }
}
