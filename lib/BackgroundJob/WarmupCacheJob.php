<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\BackgroundJob;

use OCA\Sharepoint2\Service\CacheWarmupService;
use OCP\AppFramework\Utility\ITimeFactory;
use OCP\BackgroundJob\QueuedJob;
use Psr\Log\LoggerInterface;

class WarmupCacheJob extends QueuedJob {
    public function __construct(
        ITimeFactory $time,
        private CacheWarmupService $cacheWarmupService,
        private LoggerInterface $logger,
    ) {
        parent::__construct($time);
    }

    protected function run($argument): void {
        $mountId = (int)($argument['mount_id'] ?? 0);
        if ($mountId <= 0) {
            return;
        }

        $full = (bool)($argument['full'] ?? false);
        $path = trim((string)($argument['path'] ?? ''));

        try {
            $this->cacheWarmupService->warmupMountById($mountId, $full, $path);
        } catch (\Throwable $e) {
            $this->logger->error('Sharepoint2 warmup cache job failed', [
                'mount_id' => $mountId,
                'full' => $full,
                'path' => $path,
                'error' => $e->getMessage(),
            ]);
        }
    }
}
