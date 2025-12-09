<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Storage;

use ArrayIterator;
use OC\Files\Storage\Common;
use OCP\Http\Client\IClientService;
use OCP\Constants;
use OCP\IUserSession;
use OCP\IConfig;
use OCA\Sharepoint2\Service\MSOAuth2TokenService;
use Psr\Log\LoggerInterface;
use Traversable;

class SharePointStorage extends Common {
    // Increase timeout for large folders
    private const API_TIMEOUT = 120;
    private const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

    private string $siteUrl;
    private string $libraryPath;

    private ?string $siteId = null;
    private ?string $driveId = null;

    private string $mountRootPath = '';
    private string $numericId;
    private string $clientId = '';
    private string $clientSecret = '';
    private string $tenant;

    private IClientService $httpClientService;
    private MSOAuth2TokenService $tokenService;
    private IUserSession $userSession;
    private IConfig $config;
    private LoggerInterface $logger;

    private ?string $accessToken = null;

    public function __construct(array $params) {
        $this->tokenService      = \OC::$server->get(MSOAuth2TokenService::class);
        $this->httpClientService = \OC::$server->get(IClientService::class);
        $this->userSession       = \OC::$server->get(IUserSession::class);
        $this->config            = \OC::$server->get(IConfig::class);
        $this->logger            = \OC::$server->get(LoggerInterface::class);

        $this->siteUrl     = rtrim((string)($params['site_url'] ?? ''), '/');
        $this->libraryPath = trim((string)($params['library'] ?? ''), '/');
        $this->clientId    = (string)($params['client_id'] ?? '');
        $this->clientSecret = (string)($params['client_secret'] ?? '');
        
        // Tenant Logic
        $tenantInput = trim((string)($params['tenant'] ?? ''));
        $tenantConfig = $this->config->getSystemValue('sharepoint2_tenant', '');

        if ($tenantInput !== '') {
            $this->tenant = $tenantInput;
        } elseif ($tenantConfig !== '') {
            $this->tenant = $tenantConfig;
        } else {
            $this->tenant = 'common';
        }

        $this->numericId = md5($this->siteUrl . '|' . $this->libraryPath . '|' . $this->clientId);

        parent::__construct($params);
    }

    private function log(string $message, array $context = []): void {
        $this->logger->warning('SharePointStorage: ' . $message, $context);
    }

    public function getId(): string {
        $key = implode('|', [$this->siteUrl, $this->libraryPath, $this->mountRootPath]);
        return 'sharepoint2::' . sha1($key);
    }
    
    public function test(): bool {
        if (!$this->ensureAccessToken()) return false;
        if ($this->siteUrl === '') return false;
        return $this->initialize();
    }

    private function initialize(): bool {
        if ($this->siteId !== null && $this->driveId !== null) {
            return true;
        }

        if (!$this->ensureAccessToken()) return false;

        $parts = parse_url($this->siteUrl);
        if (!is_array($parts) || empty($parts['host']) || empty($parts['path'])) {
            $this->log('initialize(): invalid siteUrl', ['siteUrl' => $this->siteUrl]);
            return false;
        }

        // 1. Get Site ID
        $site = $this->graphGet("/sites/{$parts['host']}:{$parts['path']}");
        if (!is_array($site) || empty($site['id'])) {
            $this->log('initialize(): failed to resolve siteId');
            return false;
        }
        $this->siteId = (string)$site['id'];

        // 2. Get Drive ID
        [$libraryName, $subPath] = $this->splitLibraryPath($this->libraryPath);
        
        // Loop through all drives (pagination supported)
        $allDrives = $this->fetchAllPages("/sites/{$this->siteId}/drives");
        
        foreach ($allDrives as $drive) {
            if (isset($drive['name']) && (string)$drive['name'] === $libraryName) {
                $this->driveId = (string)$drive['id'];
                break;
            }
        }

        if ($this->driveId === null) {
            $this->log('initialize(): library not found', ['lib' => $libraryName]);
            return false;
        }

        // 3. Resolve SubPath
        $this->mountRootPath = '';
        if ($subPath !== '') {
            $encodedPath = $this->encodeDrivePath($subPath);
            $item = $this->graphGet("/drives/{$this->driveId}/root:/{$encodedPath}:/");
            if (!isset($item['id'])) {
                $this->log('initialize(): subPath not found');
                return false;
            }
            $this->mountRootPath = $subPath;
        }
        return true;
    }

/**
     * Optimized List Children:
     * 1. Uses $top=999 to reduce HTTP requests by 5x.
     * 2. Uses $select to fetch only needed fields (smaller JSON).
     */
    private function listChildren(string $relativePath): array {
        if (!$this->initialize()) return [];
 
        $drivePath = $this->buildDrivePath($relativePath);

        // OPTIMIZATION: Request max page size (999) and only specific fields
        $query = '?$top=999&$select=id,name,folder,file,size,lastModifiedDateTime,eTag';

        if ($drivePath === '') {
            $graphPath = "/drives/{$this->driveId}/root/children{$query}";
        } else {
            $encoded   = $this->encodeDrivePath($drivePath);
            $graphPath = "/drives/{$this->driveId}/root:/{$encoded}:/children{$query}";
        }

        // fetchAllPages handles the pagination if > 999 items
        return $this->fetchAllPages($graphPath);
    }

    /**
     * Recursively follows @odata.nextLink to get ALL items (Pagination)
     */
    private function fetchAllPages(string $initialPath): array {
        $allItems = [];
        $nextLink = $initialPath; // Start with the relative path

        do {
            $data = $this->graphGet($nextLink);
            
            if (!is_array($data) || !isset($data['value'])) {
                break;
            }

            $allItems = array_merge($allItems, $data['value']);
            
            // Check if there is a next page
            $nextLink = $data['@odata.nextLink'] ?? null;
            
        } while ($nextLink !== null);

        return $allItems;
    }

    private function graphGet(string $pathOrUrl): ?array {
        if (!$this->ensureAccessToken()) return null;

        $client = $this->httpClientService->newClient();
        
        // Handle full URLs (from @odata.nextLink) or relative paths
        $url = str_starts_with($pathOrUrl, 'http') ? $pathOrUrl : self::GRAPH_BASE . $pathOrUrl;

        try {
            $response = $client->get($url, [
                'headers' => [
                    'Authorization' => 'Bearer ' . $this->accessToken,
                    'Accept'        => 'application/json',
                ],
                'timeout' => self::API_TIMEOUT, // Used const (120s)
            ]);

            $body = (string)$response->getBody();
            return json_decode($body, true);
        } catch (\Throwable $e) {
            $this->log('graphGet(): error', ['url' => $url, 'msg' => $e->getMessage()]);
            return null;
        }
    }

    private function ensureAccessToken(): bool {
        if ($this->accessToken !== null) return true;

        $user = $this->userSession->getUser();
        if ($user === null) {
            // CLI/Cron Fallback
            $userId = 'admin89'; 
            $this->log('ensureAccessToken(): CLI/Cron detected, using: ' . $userId);
        } else {
            $userId = $user->getUID();
        }

        $accessToken = $this->tokenService->getValidAccessToken(
            0, $userId, $this->tenant, $this->clientId, $this->clientSecret
        );

        if ($accessToken) {
            $this->accessToken = $accessToken;
            return true;
        }
        return false;
    }

    // --- Helpers ---
    private function splitLibraryPath(string $path): array {
        $parts = array_values(array_filter(explode('/', $path), 'strlen'));
        if ($parts === []) return ['Documents', ''];
        $libraryName = array_shift($parts);
        return [$libraryName, implode('/', $parts)];
    }

    private function encodeDrivePath(string $path): string {
        $parts = array_values(array_filter(explode('/', $path), 'strlen'));
        $parts = array_map('rawurlencode', $parts);
        return implode('/', $parts);
    }

    private function buildDrivePath(string $relativePath): string {
        $root = trim($this->mountRootPath, '/');
        $rel  = trim($relativePath, '/');
		
		// --- FIX: Handle "." which causes 404s ---
        if ($rel === '.') {
            $rel = '';
        }
        // -----------------------------------------
		
        if ($root === '') return $rel;
        if ($rel === '') return $root;
        return $root . '/' . $rel;
    }

	private function getItemByPath(string $path): ?array {
        if (!$this->initialize()) return null;

        // Normalize Nextcloud virtual paths to SharePoint-relative paths
        $path = trim($path, '/');
		$path = $this->normalizeLibraryPath($path);
        
        // 1. Easy check: If Nextcloud asks for root explicitly
        if ($path === '') {
            return ['id' => 'root', 'folder' => new \stdClass()];
        }

        // 2. Build the actual path on the Drive
        $drivePath = $this->buildDrivePath($path);

        // 3. CRITICAL FIX: If the result is empty (e.g. input was "."), use the ROOT endpoint.
        // DO NOT use "root:/:/" which causes the 400 Bad Request error.
        if ($drivePath === '') {
             $item = $this->graphGet("/drives/{$this->driveId}/root");
        } else {
             // Normal case: use the path-based endpoint
             $encoded = $this->encodeDrivePath($drivePath);
             $item = $this->graphGet("/drives/{$this->driveId}/root:/{$encoded}:/");
        }

        return (is_array($item) && isset($item['id'])) ? $item : null;
    }

    /**
     * Normalize a path that may include full Nextcloud VFS prefixes:
     *   <userId>/files/<mountName>/actual/sharepoint/path
     *
     * Returns only the SharePoint-relative portion:
     *   actual/sharepoint/path
     *
     * If the path does not match this pattern, returns it unchanged.
     */
    private function normalizeLibraryPath(string $path): string {
        $path = ltrim($path, '/');

        if ($path === '' || $path === '.') {
            return '';
        }

        $segments = explode('/', $path);

        // Pattern: "<userId>/files/<mountName>/..."
        if (count($segments) >= 3 && $segments[1] === 'files') {
            // Remove username, "files", and mount name
            $segments = array_slice($segments, 3);
            return implode('/', $segments);
        }

        return $path;
    }

    // --- Standard Storage Methods ---

    public function file_exists(string $path): bool {
        return $this->getItemByPath($path) !== null;
    }
    public function is_dir(string $path): bool {
        $item = $this->getItemByPath($path);
        return $item !== null && isset($item['folder']);
    }
    public function is_file(string $path): bool {
        $item = $this->getItemByPath($path);
        return $item !== null && !isset($item['folder']);
    }
    public function filetype(string $path): string {
        if ($this->is_dir($path)) return 'dir';
        if ($this->is_file($path)) return 'file';
        return '';
    }
    
    public function getDirectoryContent(string $directory = ''): Traversable {
        $children = $this->listChildren(trim($directory, '/'));
        $result = [];
        foreach ($children as $item) {
            if (!isset($item['name'])) continue;
            $isFolder = isset($item['folder']);
            $result[] = [
                'name' => (string)$item['name'],
                'size' => $isFolder ? 0 : (int)($item['size'] ?? 0),
                'mtime' => isset($item['lastModifiedDateTime']) ? strtotime((string)$item['lastModifiedDateTime']) : time(),
                'type' => $isFolder ? 'dir' : 'file',
                'mimetype' => $isFolder ? 'httpd/unix-directory' : ($item['file']['mimeType'] ?? 'application/octet-stream'),
                'permissions' => Constants::PERMISSION_READ,
                'etag' => isset($item['eTag']) ? substr(sha1((string)$item['eTag']), 0, 32) : ''
            ];
        }
        return new ArrayIterator($result);
    }

    public function fopen(string $path, string $mode) {
        if (strpbrk($mode, 'wax+') !== false) return false;
        $item = $this->getItemByPath($path);
        if (!$item || isset($item['folder'])) return false;
        
        $contentUrl = "/drives/{$this->driveId}/items/{$item['id']}/content";
        // graphGet returns array (JSON), so we need a raw download here.
        // We implement a quick raw download:
        $client = $this->httpClientService->newClient();
        try {
            $stream = fopen('php://temp', 'r+');
            $client->get(self::GRAPH_BASE . $contentUrl, [
                'headers' => ['Authorization' => 'Bearer ' . $this->accessToken],
                'timeout' => 120, // Increased timeout
                'sink' => $stream
            ]);
            rewind($stream);
            return $stream;
        } catch (\Throwable $e) { return false; }
    }
    
    // Read-only stubs
    public function mkdir(string $path): bool { return false; }
    public function rmdir(string $path): bool { return false; }
    public function unlink(string $path): bool { return false; }
    public function touch(string $path, int $mtime = null): bool { return false; }
    public function rename(string $source, string $target): bool { return false; }
    public function copy(string $source, string $target): bool { return false; }
    public function stat(string $path): array {
         // simplified stat for brevity, relying on getDirectoryContent usually
         // or you can copy your previous stat() implementation here.
         // This minimal version is safe:
         $item = $this->getItemByPath($path);
         if (!$item) return ['size'=>0, 'mtime'=>0];
         $isFolder = isset($item['folder']);
         return [
             'size' => $isFolder ? 0 : (int)($item['size']??0),
             'mtime' => time(), 
             'type' => $isFolder ? 'dir' : 'file',
             'permissions' => Constants::PERMISSION_READ
         ];
    }
    function opendir(string $path) { return false; }
}