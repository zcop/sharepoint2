<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Service;

use OCP\BackgroundJob\TimedJob;
use OCP\BackgroundJob\IJob;
use OCP\AppFramework\Utility\ITimeFactory;
use OCA\Files_External\Service\GlobalStoragesService;
use OCA\Files_External\Lib\StorageConfig;
use OCA\Sharepoint2\Service\MSOAuth2TokenService;
use OCP\IConfig;
use Psr\Log\LoggerInterface;

/**
 * Daily job to refresh all due Microsoft OAuth2 tokens.
 */
class RefreshTokensService extends TimedJob {
    private MSOAuth2TokenService $tokenService;
    private LoggerInterface $logger;
	private GlobalStoragesService $globalStoragesService;
    private IConfig $config;

    public function __construct(
        ITimeFactory $time,
        MSOAuth2TokenService $tokenService, // <--- Injects your service
        LoggerInterface $logger,
		GlobalStoragesService $globalStoragesService,
        IConfig $config
    ) {
        parent::__construct($time);

        $this->tokenService = $tokenService;
        $this->logger = $logger;
        $this->globalStoragesService = $globalStoragesService;
        $this->config = $config;

        // Run once a day (86400 seconds)
        $this->setInterval(86400);

        // Bias execution into nighttime maintenance window
        $this->setTimeSensitivity(IJob::TIME_INSENSITIVE);
    }

    /**
     * @param mixed $argument
     */
    protected function run(mixed $argument): void {
        try {
			// 1. Look for SharePoint configurations in the Database
			$configs = $this->getSharePointConfigs();
		   
			if (empty($configs)) {
                return;
            }
			
            // 2. Iterate through found configs and refresh tokens
            foreach ($configs as $config) {
                $this->tokenService->refreshAllDueTokens(
					$config['tenant'],
					$config['client_id'],
					$config['client_secret'],
                    300,
                    [$config['token_storage_id']]
                );
            }			

        } catch (\Throwable $e) {
            $this->logger->error('Sharepoint2: RefreshTokensService failed', [
                'error' => $e->getMessage(),
                'trace' => $e->getTraceAsString()
            ]);
        }
    }

    /**
     * Helper: Scans all Global External Storages to find SharePoint2 mounts.
     * Returns array of credentials found in the DB.
     */
    private function getSharePointConfigs(): array {
        $results = [];
        
        // Fetch all global storages configured in Admin UI
        $storages = $this->globalStoragesService->getStorages();

        foreach ($storages as $storage) {
            /** @var StorageConfig $storage */
            // Check if this is OUR app
            if ($storage->getBackend()->getIdentifier() === 'sharepoint2') {
                
                $options = $storage->getBackendOptions();

                // Ensure the admin actually filled out the fields
                $tenant = trim((string)($options['tenant'] ?? ''));
                if ($tenant === '') {
                    $tenant = trim((string)$this->config->getSystemValue('sharepoint2_tenant', ''));
                }
                if ($tenant === '') {
                    $tenant = 'common';
                }

                if (!empty($options['client_id']) && !empty($options['client_secret'])) {
                    $tokenKey = strtolower(rtrim((string)($options['site_url'] ?? ''), '/'))
                        . '|' . trim((string)($options['library'] ?? ''), '/')
                        . '|' . $tenant
                        . '|' . (string)$options['client_id'];
                    
                    $results[] = [
                        'tenant' => $tenant,
                        'client_id' => $options['client_id'],
                        'client_secret' => $options['client_secret'],
                        'token_storage_id' => (int)sprintf('%u', crc32($tokenKey)),
                    ];
                }
            }
        }
        return $results;
    }	
}
