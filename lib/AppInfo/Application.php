<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\AppInfo;

use OCA\Files_External\Lib\Config\IBackendProvider; // <--- Critical Import
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
        // No manual registration needed. 
        // NC32 automatically handles dependency injection for MSOAuth2TokenService.
    }

    public function boot(IBootContext $context): void {
        $context->injectFn(function (BackendService $backendService, IJobList $jobList): void {
            
            // 1. Register Backend Provider
            $backendService->registerBackendProvider($this);

            // 2. Register Cron Job
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