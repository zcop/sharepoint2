<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Storage;

use ArrayIterator;
use OC\Files\Storage\Common;
use OCP\Http\Client\IClientService;
use OCP\Constants;
use Psr\Log\LoggerInterface;
use Traversable;

class SharePointStorage extends Common {
	private const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

	private string $siteUrl;
	private string $libraryPath;

	/** @var array<string,mixed>|null */
	private ?array $token = null;

	private ?string $siteId = null;
	private ?string $driveId = null;

	/**
	 * Path inside the drive that corresponds to the mount root, relative to drive root.
	 * Example: "" (root of library) or "SubFolder/More".
	 */
	private string $mountRootPath = '';

	/**
	 * Stable id for this storage (used in oc_storages).
	 */
	private string $numericId;

	private string $clientId = '';
	private string $clientSecret = '';

	/**
	 * Tenant for token endpoint. For testing you can hardcode your tenant here
	 * (e.g. "edcvn.onmicrosoft.com" or the GUID tenant id).
	 */
	private string $tenant = 'fc153689-bfec-4013-a019-103b83f98ee1';

	private IClientService $httpClientService;

	private ?LoggerInterface $logger = null;

	/**
	 * @param array<string,mixed> $params
	 *   Expected keys from the backend config:
	 *     - site_url   (string, required)
	 *     - library    (string, required; e.g. "Documents" or "Documents/SubFolder")
	 *     - client_id  (string, required)
	 *     - client_secret (string, required)
	 *     - token      (string, encoded refresh token JSON from OAuth2 controller)
	 *     - tenant     (string, optional; tenant id or domain; defaults to "common")
	 */
	public function __construct(array $params) {
		$this->siteUrl     = rtrim((string)($params['site_url'] ?? ''), '/');
		$this->libraryPath = trim((string)($params['library'] ?? ''), '/');

		$this->clientId     = (string)($params['client_id'] ?? '');
		$this->clientSecret = (string)($params['client_secret'] ?? '');
		$this->tenant       = (string)($params['tenant'] ?? 'common');

		$this->numericId = md5($this->siteUrl . '|' . $this->libraryPath . '|' . $this->clientId);

		$server                  = \OC::$server;
		$this->httpClientService = $server->get(IClientService::class);

		// Decode token stored like files_external_onedrive:
		//   base64( gzdeflate( json_encode(tokenArray), 9 ) )
		$rawToken    = $params['token'] ?? null;
		$this->token = null;

		if (is_string($rawToken) && $rawToken !== '') {
			$compressed = base64_decode($rawToken, true);
			if ($compressed === false) {
				$this->log('token decode: base64_decode failed');
			} else {
				$json = @gzinflate($compressed);
				if ($json === false) {
					$this->log('token decode: gzinflate failed');
				} else {
					$data = json_decode($json, true);
					if (!is_array($data)) {
						$this->log('token decode: json_decode failed');
					} else {
						$this->token = $data;
					}
				}
			}
		} elseif (is_array($rawToken)) {
			$this->token = $rawToken;
		}

		parent::__construct($params);
	}

	/**
	 * Lazy logger getter.
	 */
	private function log(string $message, array $context = []): void {
		if ($this->logger === null) {
			$this->logger = \OC::$server->get(LoggerInterface::class);
		}

		$this->logger->warning('SharePointStorage: ' . $message, $context);
	}

	/**
	 * Storage id used in oc_storages (must be stable).
	 */
	public function getId(): string {
		// Build a deterministic key using stable configuration values
		$key = implode('|', [
			$this->siteUrl,
			$this->libraryPath,
			$this->mountRootPath,
		]);

		// Return a stable SharePoint2 storage ID
		return 'sharepoint2::' . sha1($key);
	}
	
	/**
	 * Nextcloud calls test() when you click the checkmark in External storages.
	 */
	public function test(): bool {
		if (!$this->ensureAccessToken()) {
			$this->log('test(): ensureAccessToken() failed');
			return false;
		}

		if ($this->siteUrl === '' || $this->libraryPath === '') {
			$this->log('test(): site_url or library is empty', [
				'siteUrl'     => $this->siteUrl,
				'libraryPath' => $this->libraryPath,
			]);
			return false;
		}

		$ok = $this->initialize();

		return $ok;
	}

	/**
	 * Resolve siteId, driveId and mountRootPath from site_url + libraryPath.
	 */
	private function initialize(): bool {
		if ($this->siteId !== null && $this->driveId !== null) {
			return true;
		}

		if (!$this->ensureAccessToken()) {
			$this->log('initialize(): ensureAccessToken() failed');
			return false;
		}

		$parts = parse_url($this->siteUrl);
		if (!is_array($parts) || empty($parts['host']) || empty($parts['path'])) {
			$this->log('initialize(): invalid siteUrl', ['siteUrl' => $this->siteUrl]);
			return false;
		}

		$host = $parts['host'];
		$path = $parts['path'];

		$site = $this->graphGet("/sites/{$host}:{$path}");
		if (!is_array($site) || empty($site['id'])) {
			$this->log('initialize(): failed to resolve siteId', [
				'host'     => $host,
				'path'     => $path,
				'response' => $site,
			]);
			return false;
		}
		$this->siteId = (string)$site['id'];

		// Split libraryPath into libraryName + subPath
		[$libraryName, $subPath] = $this->splitLibraryPath($this->libraryPath);

		// Find driveId by libraryName
		$drives = $this->graphGet("/sites/{$this->siteId}/drives");
		if (!is_array($drives) || empty($drives['value']) || !is_array($drives['value'])) {
			$this->log('initialize(): failed to list drives', ['response' => $drives]);
			return false;
		}

		$driveId = null;
		foreach ($drives['value'] as $drive) {
			if (isset($drive['name']) && (string)$drive['name'] === $libraryName) {
				$driveId = (string)$drive['id'];
				break;
			}
		}

		if ($driveId === null) {
			$this->log('initialize(): library not found among drives', [
				'libraryName' => $libraryName,
				'drives'      => $drives['value'],
			]);
			return false;
		}
		$this->driveId = $driveId;

		// Verify subPath (if any) exists as a folder
		$this->mountRootPath = '';
		if ($subPath !== '') {
			$encodedPath = $this->encodeDrivePath($subPath);
			$item        = $this->graphGet("/drives/{$this->driveId}/root:/{$encodedPath}:/");
			if (!is_array($item) || !isset($item['id']) || !isset($item['folder'])) {
				$this->log('initialize(): subPath is not a folder', [
					'subPath' => $subPath,
					'item'    => $item,
				]);
				return false;
			}
			$this->mountRootPath = $subPath;
		}
		return true;
	}

	/**
	 * Split "Documents/SubFolder" → ["Documents", "SubFolder"].
	 *
	 * @return array{0:string,1:string}
	 */
	private function splitLibraryPath(string $libraryPath): array {
		$parts = array_values(array_filter(explode('/', $libraryPath), 'strlen'));
		if ($parts === []) {
			return ['Documents', ''];
		}
		$libraryName = array_shift($parts);
		$subPath     = implode('/', $parts);

		return [$libraryName, $subPath];
	}

	/**
	 * Ensure we have a (non-expired) access token in $this->token['access_token'].
	 */
	private function ensureAccessToken(): bool {
		if ($this->token === null) {
			$this->log('ensureAccessToken(): no token structure decoded at all');
			return false;
		}

		if (!empty($this->token['access_token'])) {
			$now        = time();
			$obtainedAt = isset($this->token['obtained_at']) ? (int)$this->token['obtained_at'] : ($now - 60);
			$expiresIn  = isset($this->token['expires_in']) ? (int)$this->token['expires_in'] : 3600;

			// Refresh 5 minutes before actual expiry
			if ($now < $obtainedAt + $expiresIn - 300) {
				return true;
			}

			$this->log('ensureAccessToken(): access_token present but considered expired/near expiry');
		} else {
			$this->log('ensureAccessToken(): no access_token yet, will try refresh');
		}

		return $this->refreshAccessToken();
	}

	/**
	 * Use refresh_token + client credentials to get a fresh access_token.
	 * Does NOT persist anything to DB yet (test phase only).
	 */
	private function refreshAccessToken(): bool {
		if ($this->token === null || empty($this->token['refresh_token'])) {
			$this->log('refreshAccessToken(): no refresh_token available', [
				'has_token' => $this->token !== null,
				'keys'      => $this->token ? array_keys($this->token) : null,
			]);
			return false;
		}

		if ($this->clientId === '' || $this->clientSecret === '') {
			$this->log('refreshAccessToken(): client_id or client_secret missing', [
				'clientId_empty'     => $this->clientId === '',
				'clientSecret_empty' => $this->clientSecret === '',
			]);
			return false;
		}

		// Prefer tenant from token if present; otherwise use configured tenant (or "common").
		$tenant = isset($this->token['tenant']) && is_string($this->token['tenant'])
			? $this->token['tenant']
			: $this->tenant;

		$tokenUrl = 'https://login.microsoftonline.com/' . rawurlencode($tenant) . '/oauth2/v2.0/token';

		$scope = isset($this->token['scope']) && is_string($this->token['scope'])
			? $this->token['scope']
			: 'https://graph.microsoft.com/.default';

		$body = [
			'client_id'     => $this->clientId,
			'client_secret' => $this->clientSecret,
			'grant_type'    => 'refresh_token',
			'refresh_token' => $this->token['refresh_token'],
			'scope'         => $scope,
		];

		$client = $this->httpClientService->newClient();

		try {
			$this->log('refreshAccessToken(): POST token request', [
				'tokenUrl' => $tokenUrl,
				'scope'    => $scope,
			]);

			$response = $client->post($tokenUrl, [
				'headers' => [
					'Content-Type' => 'application/x-www-form-urlencoded',
				],
				'body'    => http_build_query($body, '', '&'),
				'timeout' => 30,
			]);

			$rawBody = (string)$response->getBody();
			$payload = json_decode($rawBody, true);

			if (!is_array($payload)) {
				$this->log('refreshAccessToken(): token endpoint returned non-JSON', [
					'bodySample' => substr($rawBody, 0, 200),
				]);
				return false;
			}

			if (!empty($payload['error'])) {
				$this->log('refreshAccessToken(): error from token endpoint', [
					'error'             => $payload['error'],
					'error_description' => $payload['error_description'] ?? null,
				]);
				return false;
			}

			if (empty($payload['access_token'])) {
				$this->log('refreshAccessToken(): no access_token in response', [
					'keys' => array_keys($payload),
				]);
				return false;
			}

			$this->token['access_token'] = $payload['access_token'];
			$this->token['expires_in']   = isset($payload['expires_in']) ? (int)$payload['expires_in'] : 3600;
			$this->token['obtained_at']  = time();

			if (!empty($payload['refresh_token'])) {
				$this->token['refresh_token'] = $payload['refresh_token'];
			}

			if (!empty($payload['scope']) && is_string($payload['scope'])) {
				$this->token['scope'] = $payload['scope'];
			}

			return true;
		} catch (\Throwable $e) {
			$this->log('refreshAccessToken(): exception', [
				'message' => $e->getMessage(),
				'class'   => get_class($e),
			]);
			return false;
		}
	}

	/**
	 * Perform Graph GET and decode JSON.
	 *
	 * @param string $path e.g. "/sites/{host}:{path}" or "/drives/{id}/root/children"
	 * @return array<string,mixed>|null
	 */
	private function graphGet(string $path): ?array {
		if (!$this->ensureAccessToken()) {
			$this->log('graphGet(): ensureAccessToken() failed', ['path' => $path]);
			return null;
		}

		$client = $this->httpClientService->newClient();

		try {
			$response = $client->get(self::GRAPH_BASE . $path, [
				'headers' => [
					'Authorization' => 'Bearer ' . $this->token['access_token'],
					'Accept'        => 'application/json',
				],
				'timeout' => 30,
			]);

			$body = (string)$response->getBody();
			$data = json_decode($body, true);

			if (!is_array($data)) {
				$this->log('graphGet(): invalid JSON', [
					'path'       => $path,
					'bodySample' => substr($body, 0, 200),
				]);
				return null;
			}

			return $data;
		} catch (\Throwable $e) {
			$this->log('graphGet(): error', [
				'path'    => $path,
				'message' => $e->getMessage(),
			]);
			return null;
		}
	}

	/**
	 * Download raw content of an item.
	 */
	private function downloadItemContent(string $itemId): ?string {
		if (!$this->ensureAccessToken()) {
			$this->log('downloadItemContent(): ensureAccessToken() failed', [
				'itemId' => $itemId,
			]);
			return null;
		}

		if ($this->driveId === null) {
			return null;
		}

		$path   = "/drives/{$this->driveId}/items/{$itemId}/content";
		$client = $this->httpClientService->newClient();

		try {
			$response = $client->get(self::GRAPH_BASE . $path, [
				'headers' => [
					'Authorization' => 'Bearer ' . $this->token['access_token'],
				],
				'timeout' => 60,
			]);
			return (string)$response->getBody();
		} catch (\Throwable $e) {
			$this->log('downloadItemContent(): error', [
				'itemId'  => $itemId,
				'message' => $e->getMessage(),
			]);
			return null;
		}
	}

	/**
	 * Encode a drive-relative path for Graph :/path:/ syntax.
	 */
	private function encodeDrivePath(string $path): string {
		$parts = array_values(array_filter(explode('/', $path), 'strlen'));
		$parts = array_map('rawurlencode', $parts);
		return implode('/', $parts);
	}

	/**
	 * Combine mountRootPath and a Nextcloud-relative path to a drive-relative path.
	 *
	 * Examples:
	 *  mountRootPath = ''   , relativePath = ''        => ''
	 *  mountRootPath = ''   , relativePath = 'foo'     => 'foo'
	 *  mountRootPath = 'A'  , relativePath = ''        => 'A'
	 *  mountRootPath = 'A'  , relativePath = 'B/C'     => 'A/B/C'
	 */
	private function buildDrivePath(string $relativePath): string {
		$root = trim($this->mountRootPath, '/');
		$rel  = trim($relativePath, '/');

		// No mount root: just use the relative path
		if ($root === '') {
			return $rel;
		}
		// Mount root only (listing root of mounted folder)
		if ($rel === '') {
			return $root;
		}
		return $root . '/' . $rel;
	}

	/**
	 * List children of a drive-relative path (relative to mount root).
	 *
	 * @param string $relativePath path relative to mount root
	 * @return array<int,array<string,mixed>>
	 */
	private function listChildren(string $relativePath): array {
		if (!$this->initialize()) {
			return [];
		}
 
		$drivePath = $this->buildDrivePath($relativePath);

		if ($drivePath === '') {
			$graphPath = "/drives/{$this->driveId}/root/children";
		} else {
			$encoded   = $this->encodeDrivePath($drivePath);
			$graphPath = "/drives/{$this->driveId}/root:/{$encoded}:/children";
		}

		$data = $this->graphGet($graphPath);

		if (!is_array($data) || !array_key_exists('value', $data) || !is_array($data['value'])) {
			$this->log('listChildren(): empty or invalid response', [
				'graphPath' => $graphPath,
				'dataType'  => gettype($data),
				'dataKeys'  => is_array($data) ? array_keys($data) : null,
			]);
			return [];
		}

		$count = count($data['value']);
		return $data['value'];
	}

	/**
	 * Get a single item by path relative to mount root.
	 *
	 * @param string $path relative path
	 * @return array<string,mixed>|null
	 */
	private function getItemByPath(string $path): ?array {
		if (!$this->initialize()) {
			return null;
		}

		$path = trim($path, '/');
		if ($path === '') {
			// Synthetic root folder
			return [
				'id'     => 'root',
				'name'   => '',
				'folder' => new \stdClass(),
			];
		}

		$drivePath = $this->buildDrivePath($path);

		$encoded = $this->encodeDrivePath($drivePath);
		$item    = $this->graphGet("/drives/{$this->driveId}/root:/{$encoded}:/");

		if (!is_array($item) || empty($item['id'])) {
			return null;
		}

		return $item;
	}

	/* ===== Implementation of required IStorage methods ===== */

	public function file_exists(string $path): bool {
		if ($path === '' || $path === '/') {
			return true;
		}

		return $this->getItemByPath($path) !== null;
	}

	public function is_dir(string $path): bool {
		if ($path === '' || $path === '/') {
			return true;
		}

		$item = $this->getItemByPath($path);
		return $item !== null && isset($item['folder']);
	}

	public function is_file(string $path): bool {
		if ($path === '' || $path === '/') {
			return false;
		}

		$item = $this->getItemByPath($path);
		return $item !== null && !isset($item['folder']);
	}

	public function filetype(string $path): string {
		if ($this->is_dir($path)) {
			return 'dir';
		}
		if ($this->is_file($path)) {
			return 'file';
		}
		return '';
	}

	public function stat(string $path): array {
		if ($path === '' || $path === '/') {
			return [
				'size'        => 0,
				'mtime'       => time(),
				'type'        => 'dir',
				'mimetype'    => 'httpd/unix-directory',
				'permissions' => Constants::PERMISSION_READ,
			];
		}

		$item = $this->getItemByPath($path);
		if ($item === null) {
			return [
				'size'        => 0,
				'mtime'       => 0,
				'type'        => '',
				'mimetype'    => '',
				'permissions' => 0,
			];
		}

		$isFolder = isset($item['folder']);
		$size     = $isFolder ? 0 : (int)($item['size'] ?? 0);
		$mtime    = isset($item['fileSystemInfo']['lastModifiedDateTime'])
			? strtotime((string)$item['fileSystemInfo']['lastModifiedDateTime'])
			: time();
		$mimetype = $isFolder
			? 'httpd/unix-directory'
			: (string)($item['file']['mimeType'] ?? 'application/octet-stream');

		return [
			'size'        => $size,
			'mtime'       => $mtime ?: time(),
			'type'        => $isFolder ? 'dir' : 'file',
			'mimetype'    => $mimetype,
			'permissions' => Constants::PERMISSION_READ,
		];
	}

	/**
	 * Report NC permissions for files/folders in this storage.
	 * For now: read-only.
	 */
	public function getPermissions(string $path): int {
		if ($this->is_dir($path) || $this->is_file($path)) {
			return Constants::PERMISSION_READ;
		}
		return 0;
	}

	/**
	 * Directory listing for Common::opendir().
	 *
	 * @param string $directory
	 * @return Traversable
	 */
	public function getDirectoryContent(string $directory = ''): Traversable {
		if (!$this->initialize()) {
			return new \ArrayIterator([]);
		}

		$relative = trim($directory, '/');

		$children = $this->listChildren($relative);

		$result = [];
			foreach ($children as $item) {
				if (!isset($item['name'])) {
					continue;
				}

				$isFolder = isset($item['folder']);
				$size     = $isFolder ? 0 : (int)($item['size'] ?? 0);
				$mtime    = isset($item['fileSystemInfo']['lastModifiedDateTime'])
					? strtotime((string)$item['fileSystemInfo']['lastModifiedDateTime'])
					: time();
				$mimetype = $isFolder
					? 'httpd/unix-directory'
					: (string)($item['file']['mimeType'] ?? 'application/octet-stream');

				// Short, DB-safe etag
				$rawEtag = isset($item['eTag']) ? (string)$item['eTag'] : '';
				$etag    = $rawEtag !== '' ? substr(sha1($rawEtag), 0, 32) : '';

				$result[] = [
					'name'        => (string)$item['name'],
					'size'        => $size,
					'mtime'       => $mtime ?: time(),
					'type'        => $isFolder ? 'dir' : 'file',
					'mimetype'    => $mimetype,
					'etag'        => $etag,
					'permissions' => \OCP\Constants::PERMISSION_READ,
				];
			}

			return new \ArrayIterator($result);
	}		


	/**
	 * Read-only fopen: downloads content into a temp stream.
	 */
	public function fopen(string $path, string $mode) {
		// Only allow read modes
		if (strpbrk($mode, 'wax+') !== false) {
			return false;
		}

		$item = $this->getItemByPath($path);
		if ($item === null || isset($item['folder']) || !isset($item['id'])) {
			return false;
		}

		$body = $this->downloadItemContent((string)$item['id']);
		if ($body === null) {
		 return false;
		}

		$stream = fopen('php://temp', 'r+');
		if ($stream === false) {
			return false;
		}

		fwrite($stream, $body);
		rewind($stream);
		return $stream;
	}

	/* ===== Read-only operations (disabled for now) ===== */

	function opendir(string $path) {
		// We let Nextcloud use getDirectoryContent() instead.
		// Returning false is fine; the important part is that the method exists.
		return false;
	}

	public function mkdir(string $path): bool {
		return false;
	}

	public function rmdir(string $path): bool {
		return false;
	}

	public function unlink(string $path): bool {
		return false;
	}

	public function touch(string $path, int $mtime = null): bool {
		return false;
	}

	public function rename(string $source, string $target): bool {
		return false;
	}

	public function copy(string $source, string $target): bool {
		return false;
	}
}
