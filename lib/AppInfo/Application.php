<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\AppInfo;

use OCA\Files_External\Event\StorageCreatedEvent;
use OCA\Files_External\Event\StorageUpdatedEvent;
use OCA\Files_External\Lib\Config\IBackendProvider;
use OCA\Files_External\Lib\Config\IAuthMechanismProvider;
use OCA\Files_External\Service\BackendService;
use OCA\Sharepoint2\Auth\OAuth2Mechanism;
use OCA\Sharepoint2\Backend\SpoBackend;
use OCA\Sharepoint2\Listener\StorageChangedListener;
use OCA\Sharepoint2\Service\RefreshTokensService;
use OCP\BackgroundJob\IJobList; 

use OCP\AppFramework\App;
use OCP\AppFramework\Bootstrap\IBootstrap;
use OCP\AppFramework\Bootstrap\IBootContext;
use OCP\AppFramework\Bootstrap\IRegistrationContext;

// MUST implement IBackendProvider to be registered as one
class Application extends App implements IBootstrap, IBackendProvider, IAuthMechanismProvider {

    public const APP_ID = 'sharepoint2';

    public function __construct(array $urlParams = []) {
        parent::__construct(self::APP_ID, $urlParams);
    }

    public function register(IRegistrationContext $context): void {
        $context->registerEventListener(StorageCreatedEvent::class, StorageChangedListener::class);
        $context->registerEventListener(StorageUpdatedEvent::class, StorageChangedListener::class);
    }

    public function boot(IBootContext $context): void {
        $context->injectFn(function (BackendService $backendService, IJobList $jobList): void {
            
            // Register backend provider once the Files External service is available
            $backendService->registerBackendProvider($this);
            $backendService->registerAuthMechanismProvider($this);

            // Guard cron job registration to avoid duplicate enqueues across boots
            if ($jobList->has(RefreshTokensService::class, null)) {
                return;
            }

            $jobList->add(RefreshTokensService::class, null);
        });
    }

    public function getBackends(): array {
        $container = $this->getContainer();
        return [
            $container->query(SpoBackend::class),
        ];
    }

    public function getAuthMechanisms(): array {
        $container = $this->getContainer();
        return [
            $container->query(OAuth2Mechanism::class),
        ];
    }
}
