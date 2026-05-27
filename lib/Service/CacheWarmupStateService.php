<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Service;

use OCP\ICache;
use OCP\ICacheFactory;
use Psr\Log\LoggerInterface;

class CacheWarmupStateService {
    private const STATE_TTL = 172800;
    private const SIGNATURE_TTL = 172800;

    private ICache $cache;

    public function __construct(
        ICacheFactory $cacheFactory,
        private LoggerInterface $logger,
    ) {
        try {
            $this->cache = $cacheFactory->createDistributed('sharepoint2_cache_state');
        } catch (\Throwable $e) {
            $this->logger->warning('Sharepoint2: distributed cache unavailable, fallback to local cache', [
                'error' => $e->getMessage(),
            ]);
            $this->cache = $cacheFactory->createLocal('sharepoint2_cache_state');
        }
    }

    public static function buildMountKey(string $siteUrl, string $library, string $tenant, string $clientId): string {
        $tokenKey = strtolower(rtrim($siteUrl, '/'))
            . '|' . trim($library, '/')
            . '|' . $tenant
            . '|' . $clientId;
        return sha1($tokenKey);
    }

    public function startWarmup(string $mountKey, int $mountId, array $level1Folders): void {
        $pending = [];
        foreach ($level1Folders as $folder) {
            $normalized = trim((string)$folder, '/');
            if ($normalized === '' || str_contains($normalized, '/')) {
                continue;
            }
            $pending[$normalized] = true;
        }

        $now = time();
        $this->cache->set($this->stateKey($mountKey), [
            'mount_id' => $mountId,
            'phase' => 'warming',
            'started_at' => $now,
            'updated_at' => $now,
            'total' => count($pending),
            'completed' => 0,
            'pending' => $pending,
            'last_error' => '',
        ], self::STATE_TTL);
    }

    public function markLevel1Done(string $mountKey, string $directory): void {
        $level1 = $this->extractLevel1($directory);
        if ($level1 === null) {
            return;
        }

        $state = $this->getWarmupState($mountKey);
        if (($state['phase'] ?? '') !== 'warming') {
            return;
        }

        $pending = is_array($state['pending'] ?? null) ? $state['pending'] : [];
        if (!isset($pending[$level1])) {
            return;
        }

        unset($pending[$level1]);
        $state['pending'] = $pending;
        $state['completed'] = ((int)($state['completed'] ?? 0)) + 1;
        $state['updated_at'] = time();

        $this->cache->set($this->stateKey($mountKey), $state, self::STATE_TTL);
    }

    public function finishWarmup(string $mountKey): void {
        $state = $this->getWarmupState($mountKey);
        $state['phase'] = 'ready';
        $state['updated_at'] = time();
        $state['pending'] = [];
        $state['last_error'] = '';
        $this->cache->set($this->stateKey($mountKey), $state, self::STATE_TTL);
    }

    public function failWarmup(string $mountKey, string $error): void {
        $state = $this->getWarmupState($mountKey);
        $state['phase'] = 'failed';
        $state['updated_at'] = time();
        $state['last_error'] = $error;
        $this->cache->set($this->stateKey($mountKey), $state, self::STATE_TTL);
    }

    public function isLevel1Pending(string $mountKey, string $directory): bool {
        $level1 = $this->extractLevel1($directory);
        if ($level1 === null) {
            return false;
        }

        $state = $this->getWarmupState($mountKey);
        if (($state['phase'] ?? '') !== 'warming') {
            return false;
        }

        $pending = is_array($state['pending'] ?? null) ? $state['pending'] : [];
        return isset($pending[$level1]);
    }

    public function getWarmupState(string $mountKey): array {
        $value = $this->cache->get($this->stateKey($mountKey));
        return is_array($value) ? $value : [];
    }

    /**
     * Returns true only when we already had a previous signature and it changed.
     */
    public function updateDirectorySignature(string $mountKey, string $directory, string $signature): bool {
        $directory = trim($directory, '/');
        $key = $this->signatureKey($mountKey, $directory);
        $existing = $this->cache->get($key);
        $old = is_array($existing) ? (string)($existing['signature'] ?? '') : '';

        $this->cache->set($key, [
            'signature' => $signature,
            'updated_at' => time(),
        ], self::SIGNATURE_TTL);

        return $old !== '' && $old !== $signature;
    }

    private function extractLevel1(string $directory): ?string {
        $normalized = trim($directory, '/');
        if ($normalized === '' || str_contains($normalized, '/')) {
            return null;
        }
        return $normalized;
    }

    private function stateKey(string $mountKey): string {
        return 'warmup:' . $mountKey;
    }

    private function signatureKey(string $mountKey, string $directory): string {
        return 'sig:' . $mountKey . ':' . sha1($directory);
    }
}
