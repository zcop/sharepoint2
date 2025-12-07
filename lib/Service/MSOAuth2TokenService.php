<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Service;

use OCP\IDBConnection;
use OCP\AppFramework\Utility\ITimeFactory;
use OCP\Http\Client\IClientService;
use OCP\DB\QueryBuilder\IQueryBuilder;

use Psr\Log\LoggerInterface;

/**
 * Central service for managing Microsoft OAuth2 tokens (Graph).
 *
 * NC 32 compatible: uses IDBConnection, ITimeFactory, IClientService, ILogger,
 * and standard QueryBuilder / PDO param types.
 */
class MSOAuth2TokenService {
	private IDBConnection $db;
	private ITimeFactory $time;
	private IClientService $httpClientService;
	private LoggerInterface $logger;

	// Table name WITHOUT prefix; NC will add prefix automatically
	private const TABLE = 'sharepoint2_tokens';

	// Safety margin before access_token expiry (seconds)
	private const EXPIRY_MARGIN = 120; // 2 minutes

	public function __construct(
		IDBConnection $db,
		ITimeFactory $time,
		IClientService $httpClientService,
		LoggerInterface $logger,
	) {
		$this->db = $db;
		$this->time = $time;
		$this->httpClientService = $httpClientService;
		$this->logger = $logger;
	}

	public function storeInitialToken(
		int $storageId,
		string $userId,
		string $tenant,
		array $tokenResponse
	): void {
		$now = $this->time->getTime();

		$accessToken  = (string)($tokenResponse['access_token']  ?? '');
		$refreshToken = (string)($tokenResponse['refresh_token'] ?? '');
		$expiresIn    = isset($tokenResponse['expires_in']) ? (int)$tokenResponse['expires_in'] : 3600;

		if ($accessToken === '' || $refreshToken === '') {
			$this->logger->warning('MSOAuth2TokenService: storeInitialToken(): missing access or refresh token', [
				'storageId' => $storageId,
				'userId'    => $userId,
			]);
			return;
		}

		$expiresAt = $now + $expiresIn;

		$qb = $this->db->getQueryBuilder();
		$qb->select('id')
			->from(self::TABLE)
			->where($qb->expr()->eq('storage_id', $qb->createNamedParameter($storageId, \PDO::PARAM_INT)))
			->andWhere($qb->expr()->eq('user_id', $qb->createNamedParameter($userId)));

		$existing = $qb->executeQuery()->fetchOne();

		if ($existing !== false) {
			// Update existing row
			$qb = $this->db->getQueryBuilder();
			$qb->update(self::TABLE)
				->set('tenant',        $qb->createNamedParameter($tenant))
				->set('access_token',  $qb->createNamedParameter($accessToken))
				->set('refresh_token', $qb->createNamedParameter($refreshToken))
				->set('expires_at',    $qb->createNamedParameter($expiresAt, \PDO::PARAM_INT))
				->set('updated_at',    $qb->createNamedParameter($now, \PDO::PARAM_INT))
				->where($qb->expr()->eq('id', $qb->createNamedParameter((int)$existing, \PDO::PARAM_INT)));

			$qb->executeStatement();
		} else {
			// Insert new row
			$qb = $this->db->getQueryBuilder();
			$qb->insert(self::TABLE)
				->setValue('storage_id',   $qb->createNamedParameter($storageId, \PDO::PARAM_INT))
				->setValue('user_id',      $qb->createNamedParameter($userId))
				->setValue('tenant',       $qb->createNamedParameter($tenant))
				->setValue('access_token', $qb->createNamedParameter($accessToken))
				->setValue('refresh_token',$qb->createNamedParameter($refreshToken))
				->setValue('expires_at',   $qb->createNamedParameter($expiresAt, \PDO::PARAM_INT))
				->setValue('created_at',   $qb->createNamedParameter($now, \PDO::PARAM_INT))
				->setValue('updated_at',   $qb->createNamedParameter($now, \PDO::PARAM_INT));

			$qb->executeStatement();
		}
	}

	public function getValidAccessToken(
		int $storageId,
		string $userId,
		string $tenant,
		string $clientId,
		string $clientSecret
	): ?string {
		$row = $this->loadTokenRow($storageId, $userId);

		if ($row === null) {
			return null;
		}

		$now = $this->time->getTime();
		$expiresAt = (int)$row['expires_at'];

		if ($now < ($expiresAt - self::EXPIRY_MARGIN)) {
			return (string)$row['access_token'];
		}

		$updated = $this->refreshTokenRow(
			$row,
			$tenant,
			$clientId,
			$clientSecret
		);

		if ($updated === null) {
			return null;
		}

		return (string)$updated['access_token'];
	}

	public function refreshAllDueTokens(
		string $tenant,
		string $clientId,
		string $clientSecret,
		int $safetyMarginSeconds = 300
	): void {
		$now = $this->time->getTime();
		$threshold = $now + $safetyMarginSeconds;

		$qb = $this->db->getQueryBuilder();
		$qb->select('*')
			->from(self::TABLE)
			->where($qb->expr()->lte('expires_at', $qb->createNamedParameter($threshold, \PDO::PARAM_INT)));

		$result = $qb->executeQuery();

		while ($row = $result->fetchAssociative()) {
			$this->refreshTokenRow($row, $tenant, $clientId, $clientSecret);
		}
		$result->free();
	}

	public function deleteTokensForStorage(int $storageId): void {
		$qb = $this->db->getQueryBuilder();
		$qb->delete(self::TABLE)
			->where($qb->expr()->eq('storage_id', $qb->createNamedParameter($storageId, \PDO::PARAM_INT)));

		$qb->executeStatement();
	}

	public function deleteTokensForUser(string $userId): void {
		$qb = $this->db->getQueryBuilder();
		$qb->delete(self::TABLE)
			->where($qb->expr()->eq('user_id', $qb->createNamedParameter($userId)));

		$qb->executeStatement();
	}

	/**
	 * @return array<string,mixed>|null
	 */
	private function loadTokenRow(int $storageId, string $userId): ?array {
		$qb = $this->db->getQueryBuilder();
		$qb->select('*')
			->from(self::TABLE)
			->where($qb->expr()->eq('storage_id', $qb->createNamedParameter($storageId,IQueryBuilder::PARAM_INT)))
			->andWhere($qb->expr()->eq('user_id', $qb->createNamedParameter($userId)))
			->setMaxResults(1);
			
		$result = $qb->executeQuery();
		$row = $result->fetch();
		$result->closeCursor();
		
		return $row ?: null;
	}

	/**
	 * @param array<string,mixed> $row
	 * @return array<string,mixed>|null
	 */
	private function refreshTokenRow(
		array $row,
		string $tenant,
		string $clientId,
		string $clientSecret
	): ?array {
		$refreshToken = (string)$row['refresh_token'];

		if ($refreshToken === '') {
			$this->logger->warning('MSOAuth2TokenService: refreshTokenRow(): missing refresh token', [
				'storageId' => $row['storage_id'] ?? null,
				'userId'    => $row['user_id'] ?? null,
			]);
			return null;
		}

		$client = $this->httpClientService->newClient();
		$now = $this->time->getTime();

		$tokenEndpoint = sprintf(
			'https://login.microsoftonline.com/%s/oauth2/v2.0/token',
			rawurlencode($tenant)
		);

		try {
			$response = $client->post($tokenEndpoint, [
				'body' => [
					'grant_type'    => 'refresh_token',
					'refresh_token' => $refreshToken,
					'client_id'     => $clientId,
					'client_secret' => $clientSecret,
					'scope'         => 'https://graph.microsoft.com/.default offline_access',
				],
			]);

			$data = json_decode((string)$response->getBody(), true);
		} catch (\Throwable $e) {
			$this->logger->warning('MSOAuth2TokenService: refreshTokenRow(): HTTP error', [
				'storageId' => $row['storage_id'] ?? null,
				'userId'    => $row['user_id'] ?? null,
				'error'     => $e->getMessage(),
			]);
			return null;
		}

		if (!is_array($data) || empty($data['access_token']) || empty($data['refresh_token'])) {
			$this->logger->warning('MSOAuth2TokenService: refreshTokenRow(): invalid response', [
				'storageId' => $row['storage_id'] ?? null,
				'userId'    => $row['user_id'] ?? null,
			]);
			return null;
		}

		$accessToken  = (string)$data['access_token'];
		$newRefresh   = (string)$data['refresh_token'];
		$expiresIn    = isset($data['expires_in']) ? (int)$data['expires_in'] : 3600;
		$expiresAt    = $now + $expiresIn;

		$qb = $this->db->getQueryBuilder();
		$qb->update(self::TABLE)
			->set('access_token',  $qb->createNamedParameter($accessToken))
			->set('refresh_token', $qb->createNamedParameter($newRefresh))
			->set('expires_at',    $qb->createNamedParameter($expiresAt, \PDO::PARAM_INT))
			->set('updated_at',    $qb->createNamedParameter($now, \PDO::PARAM_INT))
			->where($qb->expr()->eq('id', $qb->createNamedParameter((int)$row['id'], \PDO::PARAM_INT)));

		$qb->executeStatement();

		$row['access_token']  = $accessToken;
		$row['refresh_token'] = $newRefresh;
		$row['expires_at']    = $expiresAt;
		$row['updated_at']    = $now;

		return $row;
	}
}
