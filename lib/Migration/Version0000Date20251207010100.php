<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Migration;

use Closure;
use OCP\DB\ISchemaWrapper;
use OCP\Migration\IOutput;
use OCP\Migration\SimpleMigrationStep;

class Version0000Date20251207010100 extends SimpleMigrationStep {

    /**
     * @param IOutput $output
     * @param Closure $schemaClosure The \Closure returns an ISchemaWrapper
     * @param array $options
     * @return ISchemaWrapper|null
     */
    public function changeSchema(IOutput $output, Closure $schemaClosure, array $options) {
        /** @var ISchemaWrapper $schema */
        $schema = $schemaClosure();

        if (!$schema->hasTable('sharepoint2_tokens')) {
            $table = $schema->createTable('sharepoint2_tokens');

            // primary key
            $table->addColumn('id', 'integer', [
                'autoincrement' => true,
                'notnull'       => true,
            ]);

            // from your insert code:
            // ->setValue('storage_id', $qb->createNamedParameter($storageId, \PDO::PARAM_INT))
            $table->addColumn('storage_id', 'integer', [
                'notnull' => true,
                'default' => 0,
            ]);

            // ->setValue('user_id', $qb->createNamedParameter($userId))
            $table->addColumn('user_id', 'string', [
                'notnull' => true,
                'length'  => 64,
            ]);

            // ->setValue('tenant', $qb->createNamedParameter($tenant))
            $table->addColumn('tenant', 'string', [
                'notnull' => true,
                'length'  => 255,
            ]);

            // ->setValue('access_token', $qb->createNamedParameter($accessToken))
            $table->addColumn('access_token', 'text', [
                'notnull' => false,
            ]);

            // ->setValue('refresh_token', $qb->createNamedParameter($refreshToken))
            $table->addColumn('refresh_token', 'text', [
                'notnull' => false,
            ]);

            // ->setValue('expires_at', $qb->createNamedParameter($expiresAt, \PDO::PARAM_INT))
            $table->addColumn('expires_at', 'integer', [
                'notnull' => true,
                'default' => 0,
            ]);

            // ->setValue('created_at', $qb->createNamedParameter($now, \PDO::PARAM_INT))
            $table->addColumn('created_at', 'integer', [
                'notnull' => true,
                'default' => 0,
            ]);

            // ->setValue('updated_at', $qb->createNamedParameter($now, \PDO::PARAM_INT))
            $table->addColumn('updated_at', 'integer', [
                'notnull' => true,
                'default' => 0,
            ]);

            $table->setPrimaryKey(['id']);

            // One row per (storage_id, user_id, tenant) combination
            $table->addUniqueIndex(
                ['storage_id', 'user_id', 'tenant'],
                'spt2_sut_uidx'
            );

            // Helpful index if you later clean up expired tokens in a background job
            $table->addIndex(
                ['expires_at'],
                'spt2_exp_idx'
            );
        }

        return $schema;
    }
}
