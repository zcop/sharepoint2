<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Controller;

use OCP\AppFramework\Controller;
use OCP\AppFramework\Http;
use OCP\AppFramework\Http\DataResponse;
use OCP\IRequest;
use OCP\Util;
use OCP\Http\Client\IClientService;

/**
 * OAuth2 controller for SharePoint Online backend.
 *
 * Two-step flow:
 *  step=1: build authorization URL and send it to frontend
 *  step=2: exchange authorization code for tokens and return encoded token blob
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
	private const TENANT = 'fc153689-bfec-4013-a019-103b83f98ee1';

	/** Microsoft authorize endpoint (v2) for this tenant */
	private const URL_AUTHORIZE    = 'https://login.microsoftonline.com/' . self::TENANT . '/oauth2/v2.0/authorize';
	/** Microsoft token endpoint (v2) for this tenant */
	private const URL_ACCESS_TOKEN = 'https://login.microsoftonline.com/' . self::TENANT . '/oauth2/v2.0/token';
	/** Scopes we request */
	private const SCOPES           = 'offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read';

	public function __construct(string $appName, IRequest $request) {
		parent::__construct($appName, $request);
	}

	private function logError(string $message, array $context = []): void {
		if (class_exists('\OC_Log')) {
			\OC_Log::write(
				'sharepoint2',
				$message . (!empty($context) ? ' ' . json_encode($context) : ''),
				Util::ERROR
			);
		}
	}

	/**
	 * Handle OAuth2 flow.
	 *
	 * @param string|null $client_id
	 * @param string|null $client_secret
	 * @param string|null $redirect
	 * @param int|string|null $step
	 * @param string|null $code
	 *
	 * @NoAdminRequired
	 * @NoCSRFRequired
	 */
	public function receiveToken(
		?string $client_id = null,
		?string $client_secret = null,
		?string $redirect = null,
		$step = null,
		?string $code = null
	): DataResponse {
		$clientId = $client_id;
		$clientSecret = $client_secret;

		if ($clientId === null || $clientId === '' ||
			$clientSecret === null || $clientSecret === '' ||
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

		$step = (int) $step;

		// -------------------
		// STEP 1: auth URL
		// -------------------
		if ($step === 1) {
			$params = [
				'client_id'     => $clientId,
				'response_type' => 'code',
				'redirect_uri'  => $redirect,
				'response_mode' => 'query',
				'scope'         => self::SCOPES,
			];

			$authUrl = self::URL_AUTHORIZE . '?' . http_build_query($params, '', '&', PHP_QUERY_RFC3986);

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
				/** @var IClientService $clientService */
				$clientService = \OC::$server->get(IClientService::class);
				$client = $clientService->newClient();

				$body = http_build_query([
					'client_id'     => $clientId,
					'client_secret' => $clientSecret,
					'grant_type'    => 'authorization_code',
					'code'          => $code,
					'redirect_uri'  => $redirect,
					'scope'         => self::SCOPES,
				], '', '&', PHP_QUERY_RFC3986);

				$response = $client->post(self::URL_ACCESS_TOKEN, [
					'body'    => $body,
					'headers' => [
						'Content-Type' => 'application/x-www-form-urlencoded',
					],
					'timeout' => 30,
				]);

				$statusCode = $response->getStatusCode();
				$content    = (string) $response->getBody();

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
				}

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

				// ---- Build a compact token payload to store in DB ----
				$stored = [
					// keep refresh token primarily; we can always get a new access token
					'refresh_token' => $token['refresh_token'] ?? '',
					// keep access token if space allows
					'access_token'  => $token['access_token'] ?? '',
					'scope'         => $token['scope'] ?? '',
					'token_type'    => $token['token_type'] ?? '',
					'expires_in'    => isset($token['expires_in']) ? (int) $token['expires_in'] : 3600,
					'obtained_at'   => time(),
					'tenant'        => self::TENANT,
					'code_uid'      => uniqid('', true),
				];

				// First try with full data
				$encoded = base64_encode(gzdeflate(json_encode($stored), 9));

				// If still too large for oc_external_config.value (~4000), drop access_token, then scope.
				if (strlen($encoded) > 3500) {
					unset($stored['access_token']);
					$encoded = base64_encode(gzdeflate(json_encode($stored), 9));
				}
				if (strlen($encoded) > 3500) {
					unset($stored['scope']);
					$encoded = base64_encode(gzdeflate(json_encode($stored), 9));
				}

				// Final safety check
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
				$this->logError($msg);

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
