<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Storage;

use ArrayIterator;
use Icewind\Streams\CallbackWrapper;
use OC\Files\Storage\Common;
use OCP\Files\GenericFileException;
use OCP\Files\Storage\IChunkedFileWrite;
use OCP\Http\Client\IClientService;
use OCP\Constants;
use OCP\IConfig;
use OCP\ICache;
use OCP\ICacheFactory;
use OCP\ITempManager;
use OCA\Sharepoint2\Service\CacheWarmupStateService;
use OCA\Sharepoint2\Service\MSOAuth2TokenService;
use Psr\Log\LoggerInterface;
use Traversable;

class SharePointStorage extends Common implements IChunkedFileWrite {
    // Increase timeout for large folders
    private const API_TIMEOUT = 120;
    private const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
    private const APP_TOKEN_USER = '__sharepoint2_app__';
    private const WARMUP_NOTICE_FILE = '.sharepoint2-cache-building.txt';
    // Keep simple uploads conservative; larger payloads always use upload session.
    private const SIMPLE_UPLOAD_MAX_BYTES = 4194304; // 4 MiB
    private const UPLOAD_SESSION_CHUNK_BYTES = 5242880; // 5 MiB (multiple of 320 KiB)
    private const CHUNK_STAGE_DIR = '/tmp/sharepoint2-chunk-stage';
    private const RW_PERMISSIONS = Constants::PERMISSION_READ
        | Constants::PERMISSION_CREATE
        | Constants::PERMISSION_UPDATE
        | Constants::PERMISSION_DELETE;

    private string $siteUrl;
    private string $libraryPath;

    private ?string $siteId = null;
    private ?string $driveId = null;

    private string $mountRootPath = '';
    private string $clientId = '';
    private string $clientSecret = '';
    private string $tenant;
    private int $tokenStorageId;
    private string $mountCacheKey;

    private IClientService $httpClientService;
    private MSOAuth2TokenService $tokenService;
    private CacheWarmupStateService $cacheStateService;
    private IConfig $config;
    private LoggerInterface $logger;
    private ITempManager $tempManager;
    private ?ICache $localCache = null;

    private ?string $accessToken = null;
    /** @var array<string,array<string,mixed>|false> */
    private array $itemCache = [];
    /** @var array<string,array{mtime:int,etag:string}> */
    private array $directoryStateCache = [];

    public function __construct(array $params) {
        $this->tokenService      = \OC::$server->get(MSOAuth2TokenService::class);
        $this->cacheStateService = \OC::$server->get(CacheWarmupStateService::class);
        $this->httpClientService = \OC::$server->get(IClientService::class);
        $this->config            = \OC::$server->get(IConfig::class);
        $this->logger            = \OC::$server->get(LoggerInterface::class);
        $this->tempManager       = \OC::$server->get(ITempManager::class);
        try {
            $cacheFactory = \OC::$server->get(ICacheFactory::class);
            $this->localCache = $cacheFactory->createLocal('sharepoint2_mount_meta');
        } catch (\Throwable) {
            $this->localCache = null;
        }

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

        $tokenKey = strtolower($this->siteUrl) . '|' . $this->libraryPath . '|' . $this->tenant . '|' . $this->clientId;
        $this->tokenStorageId = (int)sprintf('%u', crc32($tokenKey));
        $this->mountCacheKey = sha1($tokenKey);

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
        [$libraryName, $subPath] = $this->splitLibraryPath($this->libraryPath);
        if ($this->loadCachedMountMeta()) {
            $this->mountRootPath = $subPath;
            return true;
        }

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

        $this->storeCachedMountMeta();

        // 3. Resolve SubPath
        $this->mountRootPath = '';
        if ($subPath !== '') {
            $encodedPath = $this->encodeDrivePath($subPath);
            $item = $this->graphGet("/drives/{$this->driveId}/root:/{$encodedPath}");
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
            $classification = $this->classifyGraphException($e);
            if ($classification['class'] === 'not_found') {
                return null;
            }
            $this->logClassifiedGraphError('graphGet()', $url, $e, $classification);
            return null;
        }
    }

    /**
     * @return array{status:int,class:string,retryable:bool,loggable:bool}
     */
    private function classifyGraphException(\Throwable $e): array {
        $status = $this->extractHttpStatus($e) ?? 0;
        $message = strtolower($e->getMessage());

        if ($status === 404 || str_contains($message, '404 not found') || str_contains($message, 'itemnotfound')) {
            return [
                'status' => 404,
                'class' => 'not_found',
                'retryable' => false,
                'loggable' => false,
            ];
        }
        if ($status === 401 || $status === 403) {
            return [
                'status' => $status,
                'class' => 'auth',
                'retryable' => true,
                'loggable' => true,
            ];
        }
        if ($status === 429) {
            return [
                'status' => 429,
                'class' => 'throttled',
                'retryable' => true,
                'loggable' => true,
            ];
        }
        if (in_array($status, [408, 425, 500, 502, 503, 504], true)
            || str_contains($message, 'timed out')
            || str_contains($message, 'temporarily unavailable')
            || str_contains($message, 'service unavailable')
            || str_contains($message, 'connection reset')) {
            return [
                'status' => $status,
                'class' => 'transient',
                'retryable' => true,
                'loggable' => true,
            ];
        }
        if (in_array($status, [409, 412, 423], true)) {
            return [
                'status' => $status,
                'class' => 'conflict',
                'retryable' => false,
                'loggable' => true,
            ];
        }
        if ($status >= 400 && $status <= 499) {
            return [
                'status' => $status,
                'class' => 'client',
                'retryable' => false,
                'loggable' => true,
            ];
        }
        if ($status >= 500 && $status <= 599) {
            return [
                'status' => $status,
                'class' => 'server',
                'retryable' => true,
                'loggable' => true,
            ];
        }

        return [
            'status' => $status,
            'class' => 'unknown',
            'retryable' => false,
            'loggable' => true,
        ];
    }

    /**
     * @param array{status:int,class:string,retryable:bool,loggable:bool}|null $classification
     */
    private function logClassifiedGraphError(string $op, string $url, \Throwable $e, ?array $classification = null): void {
        $classification ??= $this->classifyGraphException($e);
        if ($classification['loggable'] === false) {
            return;
        }

        $this->log($op . ': ' . $classification['class'], [
            'url' => $url,
            'status' => (string)$classification['status'],
            'retryable' => $classification['retryable'] ? '1' : '0',
            'msg' => $this->truncateErrorMessage($e->getMessage()),
        ]);
    }

    private function truncateErrorMessage(string $message, int $maxBytes = 700): string {
        if (strlen($message) <= $maxBytes) {
            return $message;
        }
        return substr($message, 0, $maxBytes) . '...';
    }

    /**
     * @param array<string,mixed> $payload
     */
    private function graphPostJson(string $path, array $payload): ?array {
        if (!$this->initialize() || !$this->ensureAccessToken()) {
            return null;
        }

        $client = $this->httpClientService->newClient();
        $url = self::GRAPH_BASE . $path;

        try {
            $response = $client->post($url, [
                'body' => json_encode($payload, JSON_THROW_ON_ERROR),
                'headers' => [
                    'Authorization' => 'Bearer ' . $this->accessToken,
                    'Accept' => 'application/json',
                    'Content-Type' => 'application/json',
                ],
                'timeout' => self::API_TIMEOUT,
            ]);
            return json_decode((string)$response->getBody(), true);
        } catch (\Throwable $e) {
            $classification = $this->classifyGraphException($e);
            if ($classification['class'] !== 'not_found') {
                $this->logClassifiedGraphError('graphPostJson()', $url, $e, $classification);
            }
            return null;
        }
    }

    /**
     * @param array<string,mixed> $payload
     */
    private function graphPatchJson(string $path, array $payload): ?array {
        if (!$this->initialize() || !$this->ensureAccessToken()) {
            return null;
        }

        $client = $this->httpClientService->newClient();
        $url = self::GRAPH_BASE . $path;

        try {
            $response = $client->patch($url, [
                'body' => json_encode($payload, JSON_THROW_ON_ERROR),
                'headers' => [
                    'Authorization' => 'Bearer ' . $this->accessToken,
                    'Accept' => 'application/json',
                    'Content-Type' => 'application/json',
                ],
                'timeout' => self::API_TIMEOUT,
            ]);
            $body = (string)$response->getBody();
            if ($body === '') {
                return [];
            }
            $decoded = json_decode($body, true);
            return is_array($decoded) ? $decoded : [];
        } catch (\Throwable $e) {
            $classification = $this->classifyGraphException($e);
            if ($classification['class'] !== 'not_found') {
                $this->logClassifiedGraphError('graphPatchJson()', $url, $e, $classification);
            }
            return null;
        }
    }

    private function graphDeletePath(string $path): bool {
        if (!$this->initialize() || !$this->ensureAccessToken()) {
            return false;
        }

        $client = $this->httpClientService->newClient();
        $url = self::GRAPH_BASE . $path;

        try {
            $client->delete($url, [
                'headers' => [
                    'Authorization' => 'Bearer ' . $this->accessToken,
                ],
                'timeout' => self::API_TIMEOUT,
            ]);
            return true;
        } catch (\Throwable $e) {
            $classification = $this->classifyGraphException($e);
            if ($classification['class'] === 'not_found') {
                return true;
            }
            $this->logClassifiedGraphError('graphDeletePath()', $url, $e, $classification);
            return false;
        }
    }

    private function uploadFileFromStream(string $path, $stream, ?int $knownSize = null): bool {
        if (!$this->initialize() || !$this->ensureAccessToken()) {
            return false;
        }

        $normalized = $this->normalizeStoragePath($path);
        if ($normalized === '') {
            return false;
        }

        $drivePath = $this->buildDrivePath($normalized);
        if ($drivePath === '') {
            return false;
        }

        if (!is_resource($stream)) {
            return false;
        }

        @rewind($stream);

        $closeWrappedStream = false;
        $size = $knownSize;
        if ($size === null) {
            $stats = fstat($stream);
            if (is_array($stats) && isset($stats['size']) && is_int($stats['size']) && $stats['size'] >= 0) {
                $size = $stats['size'];
            }
        }

        // If size is unknown, copy once to a temp stream so we can upload with a proper total length.
        if ($size === null) {
            $tmp = fopen('php://temp', 'r+');
            if ($tmp === false) {
                return false;
            }
            $copied = stream_copy_to_stream($stream, $tmp);
            if ($copied === false) {
                fclose($tmp);
                return false;
            }
            $stream = $tmp;
            $size = (int)$copied;
            $closeWrappedStream = true;
            rewind($stream);
        }

        $ok = false;
        if ($size <= self::SIMPLE_UPLOAD_MAX_BYTES) {
            $ok = $this->uploadSmallFile($drivePath, $stream);
        } else {
            $ok = $this->uploadLargeFileWithSession($drivePath, $stream, $size);
        }

        if ($closeWrappedStream) {
            fclose($stream);
        }

        if ($ok) {
            $this->invalidateCachesForMutation([$normalized, dirname($normalized)]);
        }

        return $ok;
    }

    private function uploadSmallFile(string $drivePath, $stream): bool {
        $encodedPath = $this->encodeDrivePath($drivePath);
        $url = self::GRAPH_BASE . "/drives/{$this->driveId}/root:/{$encodedPath}:/content";
        $client = $this->httpClientService->newClient();

        try {
            $client->put($url, [
                'body' => $stream,
                'headers' => [
                    'Authorization' => 'Bearer ' . $this->accessToken,
                    'Content-Type' => 'application/octet-stream',
                ],
                'timeout' => self::API_TIMEOUT,
            ]);
            return true;
        } catch (\Throwable $e) {
            $classification = $this->classifyGraphException($e);
            $this->logClassifiedGraphError('uploadSmallFile()', $url, $e, $classification);
            return false;
        }
    }

    private function uploadLargeFileWithSession(string $drivePath, $stream, int $size): bool {
        $encodedPath = $this->encodeDrivePath($drivePath);
        $session = $this->graphPostJson(
            "/drives/{$this->driveId}/root:/{$encodedPath}:/createUploadSession",
            [
                'item' => [
                    '@microsoft.graph.conflictBehavior' => 'replace',
                ],
            ]
        );

        $uploadUrl = is_array($session) ? (string)($session['uploadUrl'] ?? '') : '';
        if ($uploadUrl === '') {
            $this->log('uploadLargeFileWithSession(): missing uploadUrl', ['path' => $drivePath]);
            return false;
        }

        $client = $this->httpClientService->newClient();
        $offset = 0;
        while ($offset < $size) {
            $remaining = $size - $offset;
            $chunkSize = min(self::UPLOAD_SESSION_CHUNK_BYTES, $remaining);
            $chunk = $this->readExactChunk($stream, $chunkSize);
            if ($chunk === null || $chunk === '') {
                $this->log('uploadLargeFileWithSession(): failed to read chunk', [
                    'path' => $drivePath,
                    'offset' => $offset,
                    'requested' => $chunkSize,
                ]);
                return false;
            }

            $actualLen = strlen($chunk);
            $end = $offset + $actualLen - 1;

            if (!$this->uploadSessionChunkWithRetry($client, $uploadUrl, $chunk, $offset, $end, $size)) {
                $this->log('uploadLargeFileWithSession(): chunk upload error', [
                    'path' => $drivePath,
                    'offset' => $offset,
                    'end' => $end,
                    'size' => $size,
                ]);
                return false;
            }

            $offset += $actualLen;
        }

        if ($offset !== $size) {
            $this->log('uploadLargeFileWithSession(): incomplete upload', [
                'path' => $drivePath,
                'uploaded' => $offset,
                'expected' => $size,
            ]);
            return false;
        }

        return true;
    }

    private function readExactChunk($stream, int $targetBytes): ?string {
        if (!is_resource($stream) || $targetBytes <= 0) {
            return null;
        }

        $buffer = '';
        $idleReads = 0;
        while (strlen($buffer) < $targetBytes && !feof($stream)) {
            $piece = fread($stream, $targetBytes - strlen($buffer));
            if ($piece === false) {
                return null;
            }
            if ($piece === '') {
                // Some wrapped streams can yield short/empty interim reads.
                $idleReads++;
                if ($idleReads >= 20) {
                    break;
                }
                usleep(10000);
                continue;
            }
            $idleReads = 0;
            $buffer .= $piece;
        }

        return $buffer;
    }

    private function uploadSessionChunkWithRetry($client, string $uploadUrl, string $chunk, int $start, int $end, int $total): bool {
        $maxAttempts = 5;
        $attempt = 0;

        while ($attempt < $maxAttempts) {
            $attempt++;
            try {
                // Upload-session PUT should not include Authorization header.
                $client->put($uploadUrl, [
                    'body' => $chunk,
                    'headers' => [
                        'Content-Length' => (string)strlen($chunk),
                        'Content-Range' => "bytes {$start}-{$end}/{$total}",
                    ],
                    'timeout' => self::API_TIMEOUT,
                ]);
                return true;
            } catch (\Throwable $e) {
                $classification = $this->classifyGraphException($e);
                $retryable = $classification['retryable'];
                if (!$retryable || $attempt >= $maxAttempts) {
                    $this->log('uploadSessionChunkWithRetry(): final failure', [
                        'status' => (string)$classification['status'],
                        'class' => $classification['class'],
                        'attempt' => $attempt,
                        'start' => $start,
                        'end' => $end,
                        'msg' => $this->truncateErrorMessage($e->getMessage()),
                    ]);
                    return false;
                }

                $delayMs = 500 * (2 ** ($attempt - 1));
                if (method_exists($e, 'getResponse')) {
                    try {
                        $response = $e->getResponse();
                        if ($response !== null && method_exists($response, 'getHeaderLine')) {
                            $retryAfter = trim((string)$response->getHeaderLine('Retry-After'));
                            if ($retryAfter !== '' && ctype_digit($retryAfter)) {
                                $delayMs = max($delayMs, ((int)$retryAfter) * 1000);
                            }
                        }
                    } catch (\Throwable) {
                        // Keep default delay when response parsing fails.
                    }
                }

                usleep($delayMs * 1000);
            }
        }

        return false;
    }

    private function extractHttpStatus(\Throwable $e): ?int {
        if ((int)$e->getCode() >= 100 && (int)$e->getCode() <= 599) {
            return (int)$e->getCode();
        }
        if (method_exists($e, 'getResponse')) {
            try {
                $response = $e->getResponse();
                if ($response !== null && method_exists($response, 'getStatusCode')) {
                    $status = (int)$response->getStatusCode();
                    if ($status >= 100 && $status <= 599) {
                        return $status;
                    }
                }
            } catch (\Throwable) {
                return null;
            }
        }
        return null;
    }

    public function startChunkedWrite(string $targetPath): string {
        $normalizedTarget = $this->normalizeStoragePath($targetPath);
        if ($normalizedTarget === '') {
            throw new GenericFileException('Invalid chunked upload target path');
        }

        $token = bin2hex(random_bytes(16));
        $dir = $this->getChunkStagePath($token);
        if (!$this->ensureChunkStageRoot()) {
            throw new GenericFileException('Unable to prepare chunk staging root');
        }
        if (!@mkdir($dir, 0700, true) && !is_dir($dir)) {
            throw new GenericFileException('Unable to create chunk staging directory');
        }

        $meta = [
            'targetPath' => $normalizedTarget,
            'createdAt' => time(),
        ];
        $encoded = json_encode($meta, JSON_UNESCAPED_SLASHES);
        if (!is_string($encoded) || file_put_contents($dir . '/meta.json', $encoded, LOCK_EX) === false) {
            $this->deleteDirectoryRecursively($dir);
            throw new GenericFileException('Unable to initialize chunk staging metadata');
        }

        return $token;
    }

    public function putChunkedWritePart(string $targetPath, string $writeToken, string $chunkId, $data, ?int $size = null): ?array {
        $partId = (int)$chunkId;
        if ((string)$partId !== $chunkId || $partId < 1 || $partId > 10000) {
            throw new GenericFileException('Invalid chunk id');
        }
        if (!is_resource($data)) {
            throw new GenericFileException('Invalid chunk payload stream');
        }

        [$dir, $meta] = $this->loadChunkUploadMeta($writeToken);
        $normalizedTarget = $this->normalizeStoragePath($targetPath);
        if ($normalizedTarget === '' || !isset($meta['targetPath']) || $meta['targetPath'] !== $normalizedTarget) {
            throw new GenericFileException('Chunked upload target mismatch');
        }

        $tmpPath = sprintf('%s/%d.part.tmp', $dir, $partId);
        $finalPath = sprintf('%s/%d.part', $dir, $partId);

        $out = @fopen($tmpPath, 'wb');
        if ($out === false) {
            throw new GenericFileException('Unable to open chunk staging file');
        }

        $written = stream_copy_to_stream($data, $out);
        fclose($out);
        if ($written === false) {
            @unlink($tmpPath);
            throw new GenericFileException('Unable to stage chunk payload');
        }

        if (!@rename($tmpPath, $finalPath)) {
            @unlink($tmpPath);
            throw new GenericFileException('Unable to finalize staged chunk');
        }

        return [
            'chunkId' => $partId,
            'size' => (int)$written,
            'expected' => $size,
        ];
    }

    public function completeChunkedWrite(string $targetPath, string $writeToken): int {
        [$dir, $meta] = $this->loadChunkUploadMeta($writeToken);
        $normalizedTarget = $this->normalizeStoragePath($targetPath);
        if ($normalizedTarget === '' || !isset($meta['targetPath']) || $meta['targetPath'] !== $normalizedTarget) {
            throw new GenericFileException('Chunked upload target mismatch');
        }

        $chunkFiles = glob($dir . '/*.part');
        if (!is_array($chunkFiles) || $chunkFiles === []) {
            throw new GenericFileException('No staged chunks found');
        }

        $chunkMap = [];
        foreach ($chunkFiles as $file) {
            $base = (string)basename($file, '.part');
            if (!ctype_digit($base)) {
                continue;
            }
            $chunkMap[(int)$base] = $file;
        }
        if ($chunkMap === []) {
            throw new GenericFileException('Invalid staged chunk set');
        }

        ksort($chunkMap, SORT_NUMERIC);
        $expected = 1;
        foreach (array_keys($chunkMap) as $partId) {
            if ($partId !== $expected) {
                throw new GenericFileException('Missing chunk ' . $expected);
            }
            $expected++;
        }

        $assembled = fopen('php://temp/maxmemory:1048576', 'w+b');
        if ($assembled === false) {
            throw new GenericFileException('Unable to open assemble stream');
        }

        $totalSize = 0;
        foreach ($chunkMap as $chunkPath) {
            $in = @fopen($chunkPath, 'rb');
            if ($in === false) {
                fclose($assembled);
                throw new GenericFileException('Unable to read staged chunk');
            }

            $copied = stream_copy_to_stream($in, $assembled);
            fclose($in);
            if ($copied === false) {
                fclose($assembled);
                throw new GenericFileException('Unable to assemble staged chunks');
            }
            $totalSize += (int)$copied;
        }

        rewind($assembled);
        $uploaded = $this->uploadFileFromStream($normalizedTarget, $assembled, $totalSize);
        fclose($assembled);

        if (!$uploaded) {
            throw new GenericFileException('Unable to upload assembled file to SharePoint');
        }

        $this->deleteDirectoryRecursively($dir);
        return $totalSize;
    }

    public function cancelChunkedWrite(string $targetPath, string $writeToken): void {
        $dir = $this->getChunkStagePath($writeToken);
        if (is_dir($dir)) {
            $this->deleteDirectoryRecursively($dir);
        }
    }

    private function ensureChunkStageRoot(): bool {
        if (is_dir(self::CHUNK_STAGE_DIR)) {
            return true;
        }
        return @mkdir(self::CHUNK_STAGE_DIR, 0700, true) || is_dir(self::CHUNK_STAGE_DIR);
    }

    private function getChunkStagePath(string $writeToken): string {
        if (!preg_match('/^[a-f0-9]{32}$/', $writeToken)) {
            return self::CHUNK_STAGE_DIR . '/invalid-token';
        }
        return self::CHUNK_STAGE_DIR . '/' . $writeToken;
    }

    /**
     * @return array{0:string,1:array<string,mixed>}
     */
    private function loadChunkUploadMeta(string $writeToken): array {
        $dir = $this->getChunkStagePath($writeToken);
        if (!is_dir($dir)) {
            throw new GenericFileException('Chunk upload session not found');
        }

        $metaPath = $dir . '/meta.json';
        $raw = @file_get_contents($metaPath);
        $meta = is_string($raw) ? json_decode($raw, true) : null;
        if (!is_array($meta)) {
            throw new GenericFileException('Invalid chunk upload metadata');
        }

        return [$dir, $meta];
    }

    private function deleteDirectoryRecursively(string $path): void {
        if (!is_dir($path)) {
            return;
        }

        $entries = @scandir($path);
        if (!is_array($entries)) {
            return;
        }

        foreach ($entries as $entry) {
            if ($entry === '.' || $entry === '..') {
                continue;
            }
            $fullPath = $path . '/' . $entry;
            if (is_dir($fullPath) && !is_link($fullPath)) {
                $this->deleteDirectoryRecursively($fullPath);
            } else {
                @unlink($fullPath);
            }
        }

        @rmdir($path);
    }

    /**
     * @return array{0:string,1:string}
     */
    private function splitParentAndName(string $path): array {
        $normalized = trim($path, '/');
        if ($normalized === '') {
            return ['', ''];
        }
        $parent = trim((string)dirname($normalized), '/');
        if ($parent === '.') {
            $parent = '';
        }
        return [$parent, (string)basename($normalized)];
    }

    private function getItemIdForPath(string $path): ?string {
        $normalized = $this->normalizeStoragePath($path);
        if ($normalized === '') {
            $root = $this->graphGet("/drives/{$this->driveId}/root");
            if (is_array($root) && isset($root['id']) && $root['id'] !== '') {
                return (string)$root['id'];
            }
            return null;
        }

        $item = $this->getItemByPath($normalized);
        if (!is_array($item) || !isset($item['id']) || $item['id'] === '') {
            return null;
        }

        return (string)$item['id'];
    }

    /**
     * @param array<int,string> $paths
     */
    private function invalidateCachesForMutation(array $paths = []): void {
        $this->itemCache = [];
        $this->directoryStateCache = [];

        foreach ($paths as $path) {
            $normalized = trim($this->normalizeStoragePath($path), '/');
            if ($normalized === '') {
                continue;
            }

            $this->cacheStateService->updateDirectorySignature(
                $this->mountCacheKey,
                $normalized,
                sha1($normalized . '|' . microtime(true))
            );
        }
    }

    private function ensureAccessToken(): bool {
        if ($this->accessToken !== null) return true;

        if ($this->clientId === '' || $this->clientSecret === '' || $this->tenant === '') {
            $this->log('ensureAccessToken(): missing OAuth app configuration', [
                'hasClientId' => $this->clientId !== '',
                'hasClientSecret' => $this->clientSecret !== '',
                'hasTenant' => $this->tenant !== '',
            ]);
            return false;
        }

        $accessToken = $this->tokenService->getValidAccessToken(
            $this->tokenStorageId,
            self::APP_TOKEN_USER,
            $this->tenant,
            $this->clientId,
            $this->clientSecret
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

    private function getMountFolderName(): string {
        try {
            if (!method_exists($this, 'getMountPoint')) {
                return '';
            }
            $mountPoint = trim((string)$this->getMountPoint(), '/');
            if ($mountPoint === '') {
                return '';
            }
            return (string)basename($mountPoint);
        } catch (\Throwable) {
            return '';
        }
    }

    private function normalizeStoragePath(string $path): string {
        $normalized = trim($path, '/');
        if ($normalized === '' || $normalized === '.') {
            return '';
        }

        // Some calls (preview/wopi) may pass absolute user storage paths.
        // Convert "<user>/files/<mount>/..." into "<mount>/...".
        $wasAbsoluteStoragePath = false;
        $filesPos = strpos($normalized, '/files/');
        if ($filesPos !== false) {
            $wasAbsoluteStoragePath = true;
            $normalized = substr($normalized, $filesPos + 7);
        } elseif (str_starts_with($normalized, 'files/')) {
            $wasAbsoluteStoragePath = true;
            $normalized = substr($normalized, 6);
        }
        $normalized = trim($normalized, '/');

        // External storage APIs expect paths relative to the mount root.
        if ($wasAbsoluteStoragePath) {
            $mountFolder = $this->getMountFolderName();
            if ($mountFolder !== '' && ($normalized === $mountFolder || str_starts_with($normalized, $mountFolder . '/'))) {
                $normalized = ltrim(substr($normalized, strlen($mountFolder)), '/');
            }
        }

        return $normalized;
    }

    private function isKnownAbsentMarkerPath(string $path): bool {
        $normalized = trim($path, '/');
        if ($normalized === '') {
            return false;
        }

        $name = strtolower((string)basename($normalized));
        return $name === '.noimage' || $name === '.nomedia';
    }

    private function getItemByPath(string $path): ?array {
        if (!$this->initialize()) return null;

        // Clean up and normalize the path to mount-relative form.
        $path = $this->normalizeStoragePath($path);
        if ($this->isKnownAbsentMarkerPath($path)) {
            $this->itemCache[$path] = false;
            return null;
        }
        if (array_key_exists($path, $this->itemCache)) {
            $cached = $this->itemCache[$path];
            return $cached === false ? null : $cached;
        }
        
        // 1. Easy check: If Nextcloud asks for root explicitly
        if ($path === '') {
            $root = ['id' => 'root', 'folder' => new \stdClass()];
            $this->itemCache[$path] = $root;
            return $root;
        }

        $lookupPath = $path;
        $item = null;
        for ($attempt = 0; $attempt < 2; $attempt++) {
            // 2. Build the actual path on the Drive
            $drivePath = $this->buildDrivePath($lookupPath);

            // 3. If empty (e.g. "."), use the ROOT endpoint.
            if ($drivePath === '') {
                $item = $this->graphGet("/drives/{$this->driveId}/root");
            } else {
                $encoded = $this->encodeDrivePath($drivePath);
                $item = $this->graphGet("/drives/{$this->driveId}/root:/{$encoded}");
            }

            if (is_array($item) && isset($item['id'])) {
                if ($lookupPath !== $path) {
                    $this->itemCache[$lookupPath] = $item;
                }
                $this->itemCache[$path] = $item;
                return $item;
            }

            // Fallback once: strip the first path segment if clients leak mountpoint name.
            if ($attempt === 0 && str_contains($lookupPath, '/')) {
                $lookupPath = (string)substr($lookupPath, strpos($lookupPath, '/') + 1);
                continue;
            }
            break;
        }

        $this->itemCache[$path] = false;
        return null;
    }

    // --- Standard Storage Methods ---

    public function file_exists(string $path): bool {
        if ($this->isWarmupNoticePath($path)) {
            return true;
        }
        return $this->getItemByPath($path) !== null;
    }
    public function is_dir(string $path): bool {
        $item = $this->getItemByPath($path);
        return $item !== null && isset($item['folder']);
    }
    public function is_file(string $path): bool {
        if ($this->isWarmupNoticePath($path)) {
            return true;
        }
        $item = $this->getItemByPath($path);
        return $item !== null && !isset($item['folder']);
    }
    public function filetype(string $path): string {
        if ($this->isWarmupNoticePath($path)) {
            return 'file';
        }
        $item = $this->getItemByPath($path);
        if ($item === null) return '';
        if (isset($item['folder'])) return 'dir';
        return 'file';
    }
    
    private function cacheDirectoryListing(string $directory, array $children): void {
        $base = trim($directory, '/');
        $maxMtime = 0;
        $etagParts = [];
        foreach ($children as $item) {
            if (!is_array($item) || !isset($item['name'])) {
                continue;
            }
            $fullPath = $base === '' ? (string)$item['name'] : ($base . '/' . (string)$item['name']);
            $this->itemCache[$fullPath] = $item;
            if (isset($item['lastModifiedDateTime'])) {
                $mtime = strtotime((string)$item['lastModifiedDateTime']);
                if ($mtime !== false && $mtime > $maxMtime) {
                    $maxMtime = $mtime;
                }
            }
            $etagParts[] = (string)($item['eTag'] ?? $item['cTag'] ?? $item['id'] ?? $fullPath);
        }
        sort($etagParts);
        $signature = $this->buildLevel1Signature($children);
        $this->cacheStateService->updateDirectorySignature($this->mountCacheKey, $base, $signature);

        $this->directoryStateCache[$base] = [
            'mtime' => $maxMtime > 0 ? $maxMtime : time(),
            'etag' => substr(sha1(implode('|', $etagParts)), 0, 32),
        ];
    }

    private function buildLevel1Signature(array $children): string {
        $parts = [];
        foreach ($children as $item) {
            if (!is_array($item) || !isset($item['name'])) {
                continue;
            }

            $name = (string)$item['name'];
            $isFolder = isset($item['folder']);

            if ($isFolder) {
                // Ignore folder mtime/etag so deep (lv2/lv3) changes do not trigger lv1 rescans.
                $parts[] = 'd|' . $name . '|' . (string)($item['id'] ?? $name);
                continue;
            }

            $parts[] = 'f|' . $name
                . '|' . (string)($item['size'] ?? 0)
                . '|' . (string)($item['lastModifiedDateTime'] ?? '')
                . '|' . (string)($item['eTag'] ?? $item['cTag'] ?? '');
        }

        sort($parts);
        return sha1(implode('|', $parts));
    }

    /**
     * @return array{mtime:int,etag:string}
     */
    private function getDirectoryState(string $path): array {
        $normalized = trim($path, '/');
        if (isset($this->directoryStateCache[$normalized])) {
            return $this->directoryStateCache[$normalized];
        }

        $children = $this->listChildren($normalized);
        $this->cacheDirectoryListing($normalized, $children);
        return $this->directoryStateCache[$normalized] ?? [
            'mtime' => time(),
            'etag' => substr(sha1($normalized . '|' . time()), 0, 32),
        ];
    }

    private function getItemEtag(array $item, string $path): string {
        $raw = (string)($item['eTag'] ?? $item['cTag'] ?? $item['id'] ?? $path);
        if ($raw === '') {
            return '';
        }
        return substr(sha1($raw), 0, 32);
    }

    private function loadCachedMountMeta(): bool {
        if ($this->localCache === null) {
            return false;
        }

        $cached = $this->localCache->get($this->mountCacheKey);
        if (!is_array($cached)) {
            return false;
        }

        $siteId = $cached['siteId'] ?? null;
        $driveId = $cached['driveId'] ?? null;
        if (!is_string($siteId) || $siteId === '' || !is_string($driveId) || $driveId === '') {
            return false;
        }

        $this->siteId = $siteId;
        $this->driveId = $driveId;
        return true;
    }

    private function storeCachedMountMeta(): void {
        if ($this->localCache === null || $this->siteId === null || $this->driveId === null) {
            return;
        }

        $this->localCache->set($this->mountCacheKey, [
            'siteId' => $this->siteId,
            'driveId' => $this->driveId,
        ]);
    }
    
    public function getDirectoryContent(string $directory = ''): Traversable {
        $normalizedDirectory = $this->normalizeStoragePath($directory);
        $children = $this->listChildren($normalizedDirectory);
        $this->cacheDirectoryListing($normalizedDirectory, $children);
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
                'permissions' => $isFolder ? self::RW_PERMISSIONS : (self::RW_PERMISSIONS & ~Constants::PERMISSION_CREATE),
                'etag' => $this->getItemEtag($item, (string)$item['name'])
            ];
        }

        if ($this->shouldExposeWarmupNotice($normalizedDirectory, $result)) {
            $result[] = [
                'name' => self::WARMUP_NOTICE_FILE,
                'size' => 96,
                'mtime' => time(),
                'type' => 'file',
                'mimetype' => 'text/plain',
                'permissions' => Constants::PERMISSION_READ,
                'etag' => substr(sha1('warmup_notice|' . $normalizedDirectory), 0, 32),
            ];
        }

        return new ArrayIterator($result);
    }

    public function fopen(string $path, string $mode) {
        if ($this->isWarmupNoticePath($path)) {
            if (strpbrk($mode, 'wax+') !== false) {
                return false;
            }
            $stream = fopen('php://temp', 'r+');
            fwrite($stream, "SharePoint2 is building cache for this folder. Please wait and reload in a moment.\n");
            rewind($stream);
            return $stream;
        }

        $modeHead = strtolower($mode[0] ?? '');
        $hasPlus = str_contains($mode, '+');

        // Read-only open: needed by trashbin move-on-delete and previews.
        if ($modeHead === 'r' && !$hasPlus) {
            return $this->readStream($path);
        }

        // Append is not natively supported for this backend.
        if ($modeHead === 'a') {
            return false;
        }

        // Emulate writable stream via local temp file + writeback on close.
        if (strpbrk($mode, 'wxc+') !== false || $modeHead === 'r') {
            if ($modeHead === 'x' && $this->file_exists($path)) {
                return false;
            }

            $tmpFile = $this->createTempFile('.spt2');
            if ($tmpFile === null) {
                return false;
            }

            $needsSeed = $modeHead === 'r' || $modeHead === 'c';
            if ($needsSeed && $this->file_exists($path)) {
                $source = $this->readStream($path);
                if ($source === false) {
                    if ($modeHead === 'r') {
                        @unlink($tmpFile);
                        return false;
                    }
                } else {
                    $tmpOut = fopen($tmpFile, 'wb');
                    if ($tmpOut === false) {
                        fclose($source);
                        @unlink($tmpFile);
                        return false;
                    }
                    stream_copy_to_stream($source, $tmpOut);
                    fclose($tmpOut);
                    fclose($source);
                }
            }

            $fp = fopen($tmpFile, $mode);
            if ($fp === false) {
                @unlink($tmpFile);
                return false;
            }

            return CallbackWrapper::wrap($fp, null, null, function () use ($path, $tmpFile): void {
                try {
                    $input = @fopen($tmpFile, 'rb');
                    if ($input === false) {
                        return;
                    }
                    $ok = $this->uploadFileFromStream($path, $input);
                    fclose($input);
                    if (!$ok) {
                        $this->log('fopen(): writeback failed', ['path' => $path]);
                    }
                } finally {
                    @unlink($tmpFile);
                }
            });
        }

        return false;
    }

    private function createTempFile(string $postfix = ''): ?string {
        $tmpFile = $this->tempManager->getTemporaryFile($postfix);
        if (!is_string($tmpFile) || $tmpFile === '') {
            $this->log('createTempFile(): failed', ['postfix' => $postfix]);
            return null;
        }
        return $tmpFile;
    }

    public function readStream(string $path) {
        $item = $this->getItemByPath($path);
        if (!$item) {
            $this->log('readStream(): item lookup failed', ['path' => $path]);
            return false;
        }
        if (isset($item['folder'])) {
            $this->log('readStream(): path resolved to folder', ['path' => $path, 'itemId' => (string)($item['id'] ?? '')]);
            return false;
        }
        
        $contentUrl = "/drives/{$this->driveId}/items/{$item['id']}/content";
        $client = $this->httpClientService->newClient();
        try {
            $response = $client->get(self::GRAPH_BASE . $contentUrl, [
                'headers' => ['Authorization' => 'Bearer ' . $this->accessToken],
                'timeout' => 120, // Increased timeout
            ]);
            $stream = fopen('php://temp', 'r+');
            if ($stream === false) {
                return false;
            }

            $body = $response->getBody();
            if (is_string($body)) {
                if ($body !== '') {
                    fwrite($stream, $body);
                }
            } elseif (is_resource($body)) {
                stream_copy_to_stream($body, $stream);
            } elseif (is_object($body) && method_exists($body, 'eof') && method_exists($body, 'read')) {
                $idleReads = 0;
                while (!$body->eof()) {
                    $chunk = $body->read(8192);
                    if (!is_string($chunk)) {
                        break;
                    }
                    if ($chunk === '') {
                        $idleReads++;
                        if ($idleReads >= 20) {
                            break;
                        }
                        usleep(10000);
                        continue;
                    }
                    $idleReads = 0;
                    fwrite($stream, $chunk);
                }
            } elseif (is_object($body) && method_exists($body, 'getContents')) {
                $contents = (string)$body->getContents();
                if ($contents !== '') {
                    fwrite($stream, $contents);
                }
            } else {
                $contents = (string)$body;
                if ($contents !== '') {
                    fwrite($stream, $contents);
                }
            }
            rewind($stream);
            return $stream;
        } catch (\Throwable $e) {
            $classification = $this->classifyGraphException($e);
            if ($classification['class'] !== 'not_found') {
                $this->log('readStream(): content download error', [
                    'class' => $classification['class'],
                    'status' => (string)$classification['status'],
                    'retryable' => $classification['retryable'] ? '1' : '0',
                    'path' => $path,
                    'itemId' => (string)($item['id'] ?? ''),
                    'url' => self::GRAPH_BASE . $contentUrl,
                    'msg' => $this->truncateErrorMessage($e->getMessage()),
                ]);
            }
            return false;
        }
    }
    
    public function mkdir(string $path): bool {
        $normalized = $this->normalizeStoragePath($path);
        if ($normalized === '' || $this->file_exists($normalized)) {
            return false;
        }

        [$parent, $name] = $this->splitParentAndName($normalized);
        if ($name === '') {
            return false;
        }

        $driveParent = $this->buildDrivePath($parent);
        if ($parent !== '' && $this->getItemByPath($parent) === null) {
            return false;
        }

        $endpoint = $driveParent === ''
            ? "/drives/{$this->driveId}/root/children"
            : "/drives/{$this->driveId}/root:/{$this->encodeDrivePath($driveParent)}:/children";

        $created = $this->graphPostJson($endpoint, [
            'name' => $name,
            'folder' => new \stdClass(),
            '@microsoft.graph.conflictBehavior' => 'fail',
        ]);

        if (!is_array($created) || !isset($created['id'])) {
            return false;
        }

        $this->invalidateCachesForMutation([$normalized, $parent]);
        return true;
    }

    public function rmdir(string $path): bool {
        $normalized = $this->normalizeStoragePath($path);
        if ($normalized === '') {
            return false;
        }

        $item = $this->getItemByPath($normalized);
        if (!is_array($item) || !isset($item['folder']) || !isset($item['id'])) {
            return false;
        }

        $deleted = $this->graphDeletePath("/drives/{$this->driveId}/items/{$item['id']}");
        if ($deleted) {
            $this->invalidateCachesForMutation([$normalized, dirname($normalized)]);
        }
        return $deleted;
    }

    public function unlink(string $path): bool {
        $normalized = $this->normalizeStoragePath($path);
        if ($normalized === '') {
            return false;
        }

        $item = $this->getItemByPath($normalized);
        if (!is_array($item) || isset($item['folder']) || !isset($item['id'])) {
            return false;
        }

        $deleted = $this->graphDeletePath("/drives/{$this->driveId}/items/{$item['id']}");
        if ($deleted) {
            $this->invalidateCachesForMutation([$normalized, dirname($normalized)]);
        }
        return $deleted;
    }

    public function touch(string $path, ?int $mtime = null): bool {
        $normalized = $this->normalizeStoragePath($path);
        if ($normalized === '') {
            return false;
        }

        if ($this->is_dir($normalized)) {
            return true;
        }

        if ($this->file_exists($normalized)) {
            $item = $this->getItemByPath($normalized);
            if (!is_array($item) || !isset($item['id'])) {
                return false;
            }

            $dt = gmdate('Y-m-d\TH:i:s\Z', $mtime ?? time());
            $patched = $this->graphPatchJson("/drives/{$this->driveId}/items/{$item['id']}", [
                'fileSystemInfo' => [
                    'lastModifiedDateTime' => $dt,
                ],
            ]);
            if ($patched === null) {
                return false;
            }

            $this->invalidateCachesForMutation([$normalized, dirname($normalized)]);
            return true;
        }

        $stream = fopen('php://temp', 'r+');
        if ($stream === false) {
            return false;
        }
        $ok = $this->uploadFileFromStream($normalized, $stream);
        fclose($stream);
        return $ok;
    }

    public function rename(string $source, string $target): bool {
        $sourcePath = $this->normalizeStoragePath($source);
        $targetPath = $this->normalizeStoragePath($target);

        if ($sourcePath === '' || $targetPath === '' || $sourcePath === $targetPath) {
            return false;
        }

        $sourceItem = $this->getItemByPath($sourcePath);
        if (!is_array($sourceItem) || !isset($sourceItem['id'])) {
            return false;
        }

        [$targetParent, $targetName] = $this->splitParentAndName($targetPath);
        if ($targetName === '') {
            return false;
        }

        $targetParentId = $this->getItemIdForPath($targetParent);
        if ($targetParentId === null) {
            return false;
        }

        if ($this->file_exists($targetPath)) {
            if ($this->is_dir($targetPath)) {
                if (!$this->rmdir($targetPath)) {
                    return false;
                }
            } elseif (!$this->unlink($targetPath)) {
                return false;
            }
        }

        $patched = $this->graphPatchJson("/drives/{$this->driveId}/items/{$sourceItem['id']}", [
            'name' => $targetName,
            'parentReference' => [
                'id' => $targetParentId,
            ],
        ]);

        if ($patched === null) {
            return false;
        }

        $this->invalidateCachesForMutation([$sourcePath, dirname($sourcePath), $targetPath, $targetParent]);
        return true;
    }

    public function copy(string $source, string $target): bool {
        $sourcePath = $this->normalizeStoragePath($source);
        $targetPath = $this->normalizeStoragePath($target);
        if ($sourcePath === '' || $targetPath === '') {
            return false;
        }

        if ($this->is_dir($sourcePath)) {
            if (!$this->mkdir($targetPath)) {
                return false;
            }
            $children = $this->listChildren($sourcePath);
            foreach ($children as $child) {
                if (!isset($child['name'])) {
                    continue;
                }
                $name = (string)$child['name'];
                if (!$this->copy($sourcePath . '/' . $name, $targetPath . '/' . $name)) {
                    return false;
                }
            }
            return true;
        }

        $sourceStream = $this->fopen($sourcePath, 'r');
        if ($sourceStream === false) {
            return false;
        }
        $ok = $this->uploadFileFromStream($targetPath, $sourceStream);
        fclose($sourceStream);
        return $ok;
    }

    public function file_put_contents(string $path, mixed $data): int|float|false {
        if (is_resource($data)) {
            $stats = fstat($data);
            $size = null;
            if (is_array($stats) && isset($stats['size']) && is_int($stats['size']) && $stats['size'] >= 0) {
                $size = $stats['size'];
            }
            $ok = $this->uploadFileFromStream($path, $data, $size);
            if (!$ok) {
                return false;
            }
            return $size ?? 0;
        }

        $stream = fopen('php://temp', 'r+');
        if ($stream === false) {
            return false;
        }

        $stringData = (string)$data;
        $bytes = fwrite($stream, $stringData);
        if ($bytes === false) {
            fclose($stream);
            return false;
        }
        rewind($stream);

        $ok = $this->uploadFileFromStream($path, $stream, strlen($stringData));
        $size = fstat($stream)['size'] ?? 0;
        fclose($stream);

        return $ok ? (int)$size : false;
    }

    public function needsPartFile(): bool {
        // Force Nextcloud to keep temporary ".part" assembly local and only stream final content to SharePoint.
        return false;
    }

    public function writeStream(string $path, $stream, ?int $size = null): int {
        if (!is_resource($stream)) {
            return 0;
        }

        $ok = $this->uploadFileFromStream($path, $stream, $size);
        if (!$ok) {
            return 0;
        }

        if ($size !== null) {
            return $size;
        }

        $stats = fstat($stream);
        if (is_array($stats) && isset($stats['size']) && is_int($stats['size']) && $stats['size'] >= 0) {
            return $stats['size'];
        }

        return 0;
    }
    public function stat(string $path): array {
         if ($this->isWarmupNoticePath($path)) {
             return [
                 'size' => 96,
                 'mtime' => time(),
                 'type' => 'file',
                 'permissions' => Constants::PERMISSION_READ,
                 'etag' => substr(sha1('warmup_notice|' . trim($path, '/')), 0, 32),
             ];
         }

         $item = $this->getItemByPath($path);
         if (!$item) return ['size'=>0, 'mtime'=>0];
         $isFolder = isset($item['folder']);
         $mtime = isset($item['lastModifiedDateTime']) ? strtotime((string)$item['lastModifiedDateTime']) : time();
         $etag = $this->getItemEtag($item, $path);

         if ($isFolder) {
             $state = $this->getDirectoryState($path);
             $mtime = max((int)$mtime, $state['mtime']);
             $etag = $state['etag'];
         }

         return [
             'size' => $isFolder ? 0 : (int)($item['size'] ?? 0),
             'mtime' => (int)$mtime,
             'type' => $isFolder ? 'dir' : 'file',
             'permissions' => $isFolder ? self::RW_PERMISSIONS : (self::RW_PERMISSIONS & ~Constants::PERMISSION_CREATE),
             'etag' => $etag,
         ];
    }

    public function hasUpdated(string $path, int $time): bool {
        $normalized = $this->normalizeStoragePath($path);
        $item = $this->getItemByPath($normalized);
        if ($item === null || !isset($item['folder'])) {
            return parent::hasUpdated($path, $time);
        }

        $children = $this->listChildren($normalized);
        $signature = $this->buildLevel1Signature($children);
        $signatureChanged = $this->cacheStateService->updateDirectorySignature($this->mountCacheKey, $normalized, $signature);

        if ($signatureChanged) {
            return true;
        }

        return parent::hasUpdated($path, $time);
    }

    private function isWarmupNoticePath(string $path): bool {
        $normalized = trim($path, '/');
        if ($normalized === '' || !str_ends_with($normalized, '/' . self::WARMUP_NOTICE_FILE)) {
            return false;
        }

        $parent = trim(dirname($normalized), '/');
        return $this->cacheStateService->isLevel1Pending($this->mountCacheKey, $parent);
    }

    /**
     * @param array<int,array<string,mixed>> $entries
     */
    private function shouldExposeWarmupNotice(string $directory, array $entries): bool {
        if ($entries !== []) {
            return false;
        }
        if (!$this->cacheStateService->isLevel1Pending($this->mountCacheKey, $directory)) {
            return false;
        }
        if ($this->isScannerContext()) {
            return false;
        }
        return true;
    }

    private function isScannerContext(): bool {
        foreach (debug_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 12) as $frame) {
            if (($frame['class'] ?? '') === 'OC\\Files\\Cache\\Scanner') {
                return true;
            }
        }
        return false;
    }

    function opendir(string $path) { return false; }
}
