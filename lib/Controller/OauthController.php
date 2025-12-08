<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Controller;

use OCA\Sharepoint2\Service\MSOAuth2TokenService;
use OCP\AppFramework\Controller;
use OCP\AppFramework\Http;
use OCP\AppFramework\Http\DataResponse;
use OCP\Http\Client\IClientService;
use OCP\IRequest;
use OCP\IUserSession;
use OCP\IConfig;
use Psr\Log\LoggerInterface;

/**
 * OAuth2 controller for SharePoint Online backend.
 *
 * Two-step flow:
 *  step=1: build authorization URL and send it to frontend
 *  step=2: exchange authorization code for tokens, store them via MSOAuth2TokenService
 *
 * Frontend logic lives in js/sharepoint2.js.
 */
class OauthController extends Controller {

	/**
	 * IMPORTANT: set this to your tenant ID (GUID) or tenant domain.
	 *
	 * Examples:
	 *  - 'edcvn.onmicrosoft.com'
	 *  - 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
	 */
	private const TENANT_FALLBACK = 'common';
	
	private const SCOPES           = 'offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read';

	private IClientService $clientService;
	private MSOAuth2TokenService $tokenService;
	private IUserSession $userSession;
	private IConfig $config;
	private LoggerInterface $logger;

	public function __construct(
		string $appName,
		IRequest $request,
		IClientService $clientService,
		MSOAuth2TokenService $tokenService,
		IUserSession $userSession,
		IConfig $config,
		LoggerInterface $logger
	) {
		parent::__construct($appName, $request);
		$this->clientService = $clientService;
		$this->tokenService  = $tokenService;
		$this->userSession   = $userSession;
		$this->config        = $config;
		$this->logger        = $logger;
	}

	private function logError(string $message, array $context = []): void {
		$context['app'] = 'sharepoint2';
		$this->logger->error($message, $context);
	}

	/**
     * Helper to resolve Tenant ID order:
     * 1. JS Input (param)
     * 2. config.php (system value)
     * 3. Fallback ('common' or your test ID)
     */
    private function resolveTenant(?string $input): string {
        if ($input !== null && $input !== '') {
            return $input;
        }
        $systemValue = $this->config->getSystemValue('sharepoint2_tenant', '');
        if ($systemValue !== '') {
            return $systemValue;
        }
        return self::TENANT_FALLBACK;
    }
	
	/**
     * Handle OAuth2 flow.
     * Added $tenant parameter to accept input from JS.
     * * @NoAdminRequired
     * @NoCSRFRequired
     */
	public function receiveToken(
		?string $client_id = null,
		?string $client_secret = null,
		?string $tenant = null,
		?string $redirect = null,
		$step = null,
		?string $code = null
	): DataResponse {
		$resolvedTenant = $this->resolveTenant($tenant);

		if ($client_id === null || $client_id === '' ||
			$client_secret === null || $client_secret === '' ||
			$redirect === null || $redirect === '') {

			$msg = 'Missing client_id, client_secret or redirect parameter';
			$this->logError($msg, ['step' => $step]);

			return new DataResponse(
				[
					'status' => 'error',
					'data'   => ['message' => $msg],
				],
				Http::STATUS_OK
			);
		}

		if ($step === null) {
			$msg = 'Missing step parameter';
			$this->logError($msg);

			return new DataResponse(
				[
					'status' => 'error',
					'data'   => ['message' => $msg],
				],
				Http::STATUS_OK
			);
		}

		$step = (int)$step;

		// -------------------
		// STEP 1: auth URL
		// -------------------
		if ($step === 1) {
			// Build URL dynamically using the resolved tenant
			$baseUrl = 'https://login.microsoftonline.com/' . $resolvedTenant . '/oauth2/v2.0/authorize';
			$params = [
				'client_id'     => $client_id,
				'response_type' => 'code',
				'redirect_uri'  => $redirect,
				'response_mode' => 'query',
				'scope'         => self::SCOPES,
			];

			$authUrl = $baseUrl . '?' . http_build_query($params, '', '&', PHP_QUERY_RFC3986);

			return new DataResponse(
				[
					'status' => 'success',
					'data'   => ['url' => $authUrl],
				],
				Http::STATUS_OK
			);
		}

		// -------------------
		// STEP 2: token exchange
		// -------------------
		if ($step === 2) {
			if ($code === null || $code === '') {
				$msg = 'Missing authorization code';
				$this->logError($msg);

				return new DataResponse(
					[
						'status' => 'error',
						'data'   => ['message' => $msg],
					],
					Http::STATUS_OK
				);
			}

			try {
				$user = $this->userSession->getUser();
				if ($user === null) {
					$msg = 'No logged-in user during OAuth2 token exchange';
					$this->logError($msg);

					return new DataResponse(
						[
							'status' => 'error',
							'data'   => ['message' => $msg],
						],
						Http::STATUS_OK
					);
				}
				$userId = $user->getUID();

				$client = $this->clientService->newClient();

				// Build Token URL dynamically
				$tokenUrl = 'https://login.microsoftonline.com/' . $resolvedTenant . '/oauth2/v2.0/token';
				
				$body = http_build_query([
					'client_id'     => $client_id,
					'client_secret' => $client_secret,
					'grant_type'    => 'authorization_code',
					'code'          => $code,
					'redirect_uri'  => $redirect,
					'scope'         => self::SCOPES,
				], '', '&', PHP_QUERY_RFC3986);

				$response = $client->post($tokenUrl, [
					'body'    => $body,
					'headers' => [
						'Content-Type' => 'application/x-www-form-urlencoded',
					],
					'timeout' => 30,
				]);

				//$statusCode = $response->getStatusCode();
				$content    = (string)$response->getBody();
				/*
				if ($statusCode < 200 || $statusCode >= 300) {
					$msg = 'Token endpoint returned HTTP ' . $statusCode;
					$this->logError($msg, ['body' => $content]);

					return new DataResponse(
						[
							'status' => 'error',
							'data'   => [
								'message' => $msg,
								'body'    => $content,
							],
						],
						Http::STATUS_OK
					);
				}*/

				$token = json_decode($content, true);
				if (!\is_array($token) || !isset($token['access_token'])) {
					$msg = 'Invalid token response from Microsoft';
					$this->logError($msg, ['body' => $content]);

					return new DataResponse(
						[
							'status' => 'error',
							'data'   => [
								'message' => $msg,
								'body'    => $content,
							],
						],
						Http::STATUS_OK
					);
				}

				// Store token in DB using central service.
				// For now we use storageId = 0 so the token is per (user, tenant).
				$this->tokenService->storeInitialToken(
					0,
					$userId,
					$resolvedTenant,
					$token
				);

				// Build a compact token payload to store in external config as legacy field.
				// This keeps JS code unchanged, but SharePointStorage will ignore it
				// and always use MSOAuth2TokenService instead.
				$stored = [
					'refresh_token' => $token['refresh_token'] ?? '',
					'access_token'  => $token['access_token'] ?? '',
					'scope'         => $token['scope'] ?? '',
					'token_type'    => $token['token_type'] ?? '',
					'expires_in'    => isset($token['expires_in']) ? (int)$token['expires_in'] : 3600,
					'obtained_at'   => time(),
					'tenant'        => $resolvedTenant,
					'code_uid'      => uniqid('', true),
				];

				$encoded = base64_encode(gzdeflate(json_encode($stored), 9));

				if (strlen($encoded) > 3500) {
					unset($stored['access_token']);
					$encoded = base64_encode(gzdeflate(json_encode($stored), 9));
				}
				if (strlen($encoded) > 3500) {
					unset($stored['scope']);
					$encoded = base64_encode(gzdeflate(json_encode($stored), 9));
				}

				if (strlen($encoded) > 3900) {
					$msg = 'Encoded token still too large to store (length=' . strlen($encoded) . ')';
					$this->logError($msg);

					return new DataResponse(
						[
							'status' => 'error',
							'data'   => ['message' => $msg],
						],
						Http::STATUS_OK
					);
				}

				return new DataResponse(
					[
						'status' => 'success',
						'data'   => ['token' => $encoded],
					],
					Http::STATUS_OK
				);
			} catch (\Throwable $e) {
				$msg = 'Exception during token exchange: ' . $e->getMessage();
				$this->logError($msg, ['exception' => $e]);

				return new DataResponse(
					[
						'status' => 'error',
						'data'   => ['message' => $msg],
					],
					Http::STATUS_OK
				);
			}
		}

		// Unknown step
		$msg = 'Invalid step parameter';
		$this->logError($msg, ['step' => $step]);

		return new DataResponse(
			[
				'status' => 'error',
				'data'   => ['message' => $msg],
			],
			Http::STATUS_OK
		);
	}
}
