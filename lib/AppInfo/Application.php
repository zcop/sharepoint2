<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\AppInfo;

use OCA\Files_External\Lib\Config\IBackendProvider;
use OCA\Files_External\Service\BackendService;
use OCA\Sharepoint2\Backend\SpoBackend;
use OCA\Sharepoint2\Service\RefreshTokensService;
use OCP\BackgroundJob\IJobList; 

use OCP\AppFramework\App;
use OCP\AppFramework\Bootstrap\IBootstrap;
use OCP\AppFramework\Bootstrap\IBootContext;
use OCP\AppFramework\Bootstrap\IRegistrationContext;

// MUST implement IBackendProvider to be registered as one
class Application extends App implements IBootstrap, IBackendProvider {

    public const APP_ID = 'sharepoint2';

    public function __construct(array $urlParams = []) {
        parent::__construct(self::APP_ID, $urlParams);
    }

    public function register(IRegistrationContext $context): void {
        // No container tweaks needed at registration time.
    }

    public function boot(IBootContext $context): void {
        $context->injectFn(function (BackendService $backendService, IJobList $jobList): void {
            
            // Register backend provider once the Files External service is available
            $backendService->registerBackendProvider($this);

            // Guard cron job registration to avoid duplicate enqueues across boots
            if ($jobList->has(RefreshTokensService::class)) {
                return;
            }

            $jobList->add(RefreshTokensService::class);
        });
    }

    public function getBackends(): array {
        $container = $this->getContainer();
        return [
            $container->query(SpoBackend::class),
        ];
    }
}
