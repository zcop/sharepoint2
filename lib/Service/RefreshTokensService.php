<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Service;

use OCP\BackgroundJob\TimedJob;
use OCP\BackgroundJob\IJob;
use OCP\AppFramework\Utility\ITimeFactory;
use OCP\IConfig; 
use OCA\Sharepoint2\Service\MSOAuth2TokenService;
use Psr\Log\LoggerInterface;

/**
 * Daily job to refresh all due Microsoft OAuth2 tokens.
 */
class RefreshTokensService extends TimedJob {
    private MSOAuth2TokenService $tokenService;
    private LoggerInterface $logger;
    private IConfig $config;

    public function __construct(
        ITimeFactory $time,
        MSOAuth2TokenService $tokenService, // <--- Injects your service
        LoggerInterface $logger,
        IConfig $config // <--- Injects Config (Replaces \OC::$server)
    ) {
        parent::__construct($time);

        $this->tokenService = $tokenService;
        $this->logger = $logger;
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
            // 1. Get credentials from System Config
            // We fetch these here (runtime) rather than constructor to ensure 
            // we always have the latest values from config.php
            $tenant = (string) $this->config->getSystemValue('sharepoint2_tenant', '');
            $clientId = (string) $this->config->getSystemValue('sharepoint2_client_id', '');
            $clientSecret = (string) $this->config->getSystemValue('sharepoint2_client_secret', '');

            // 2. Validate Credentials
            if ($tenant === '' || $clientId === '' || $clientSecret === '') {
                $this->logger->warning('Sharepoint2: RefreshTokensService cannot run. Missing tenant, clientId, or clientSecret in config.php.');
                return;
            }

            // 3. Call your Service
            // Your service method signature is:
            // refreshAllDueTokens(string $tenant, string $clientId, string $clientSecret, int $safetyMarginSeconds = 300)
            $this->tokenService->refreshAllDueTokens(
                $tenant,
                $clientId,
                $clientSecret,
                300 // Refresh if expiring within 5 minutes
            );

        } catch (\Throwable $e) {
            $this->logger->error('Sharepoint2: RefreshTokensService failed', [
                'error' => $e->getMessage(),
                'trace' => $e->getTraceAsString()
            ]);
        }
    }
}