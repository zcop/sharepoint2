<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\AppInfo;

use OCA\Files_External\Lib\Config\IBackendProvider;
use OCA\Files_External\Service\BackendService;
use OCA\Sharepoint2\Backend\SpoBackend;
use OCP\AppFramework\App;
use OCP\AppFramework\Bootstrap\IBootstrap;
use OCP\AppFramework\Bootstrap\IBootContext;
use OCP\AppFramework\Bootstrap\IRegistrationContext;

class Application extends App implements IBootstrap, IBackendProvider {

    public const APP_ID = 'sharepoint2';

    public function __construct(array $urlParams = []) {
        parent::__construct(self::APP_ID, $urlParams);
    }

    public function register(IRegistrationContext $context): void {
        // No special services yet
    }

    public function boot(IBootContext $context): void {
        $context->injectFn(function (BackendService $backendService): void {
            // Register THIS object as backend provider
            $backendService->registerBackendProvider($this);
        });
    }

    /**
     * Return all external-storage backends this app provides.
     * For now, just a single dummy SharePoint Online backend.
     */
     public function getBackends(): array {
    	$container = $this->getContainer();

    	return [
        	$container->query(\OCA\Sharepoint2\Backend\SpoBackend::class),
    	];
     }
}
