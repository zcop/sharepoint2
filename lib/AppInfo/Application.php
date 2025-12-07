<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\AppInfo;

use OCA\Files_External\Lib\Config\IBackendProvider;
use OCA\Files_External\Service\BackendService;
use OCA\Sharepoint2\Backend\SpoBackend;
use OCA\Sharepoint2\Service\MSOAuth2TokenService;

use OCP\AppFramework\App;
use OCP\AppFramework\Bootstrap\IBootstrap;
use OCP\AppFramework\Bootstrap\IBootContext;
use OCP\AppFramework\Bootstrap\IRegistrationContext;

use OCP\IDBConnection;
use OCP\AppFramework\Utility\ITimeFactory;
use OCP\Http\Client\IClientService;
use Psr\Log\LoggerInterface;

class Application extends App implements IBootstrap, IBackendProvider {

	public const APP_ID = 'sharepoint2';

	public function __construct(array $urlParams = []) {
		parent::__construct(self::APP_ID, $urlParams);
	}

	/**
	 * Define services for DI container
	 */
	public function register(IRegistrationContext $context): void {
		// Register MSOAuth2TokenService so it can be injected anywhere
		$context->registerService(MSOAuth2TokenService::class, function (): MSOAuth2TokenService {
			$server = \OC::$server;

			return new MSOAuth2TokenService(
				$server->get(IDBConnection::class),
				$server->get(ITimeFactory::class),
				$server->get(IClientService::class),
				$server->get(LoggerInterface::class)
			);
		});
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
			$container->query(SpoBackend::class),
		];
	}
}
