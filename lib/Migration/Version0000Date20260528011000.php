<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Migration;

use Closure;
use OCP\DB\ISchemaWrapper;
use OCP\Migration\IOutput;
use OCP\Migration\SimpleMigrationStep;

class Version0000Date20260528011000 extends SimpleMigrationStep {

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
            return $schema;
        }

        $table = $schema->getTable('sharepoint2_tokens');
        if (!$table->hasIndex('spt2_user_idx')) {
            $table->addIndex(['user_id'], 'spt2_user_idx');
        }

        return $schema;
    }
}
