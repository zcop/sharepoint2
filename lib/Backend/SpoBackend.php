<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Backend;

use OCA\Files_External\Lib\Backend\Backend;
use OCA\Files_External\Lib\Auth\AuthMechanism;
use OCA\Files_External\Lib\DefinitionParameter;
use OCP\IL10N;

/**
 * SharePoint Online (OAuth2) backend definition.
 */
class SpoBackend extends Backend {

    public function __construct(IL10N $l10n) {
		// Temporary hardcoded path until you add app.svg
        $appWebPath = 'apps/sharepoint2';

        $this
            // Backend identifier (also becomes CSS class on row)
            ->setIdentifier('sharepoint2')

            // Still using Local as a placeholder storage implementation
            ->setStorageClass('\\OCA\\Sharepoint2\\Storage\\SharePointStorage')

            // Label in "Add storage" dropdown
            ->setText('SharePoint Online (OAuth2)')

            // Use generic OAuth2 auth mechanism
            ->addAuthScheme(AuthMechanism::SCHEME_OAUTH2)

            // Load our custom JS
            ->addCustomJs("../../../$appWebPath/js/sharepoint2")

            ->setPriority(200)

            // ---------- custom parameters ----------
            ->addParameters([
                // Required: which SharePoint site
                new DefinitionParameter(
                    'site_url',
                    $l10n->t('SharePoint Site URL (e.g. https://tenant.sharepoint.com/sites/MySite)')
                ),

                // Library path, e.g. "Documents" or "Documents/SubFolder"
                (new DefinitionParameter(
                    'library',
                    $l10n->t('Library path (e.g. Documents or Documents/SubFolder)')
                ))
                    ->setFlag(DefinitionParameter::FLAG_OPTIONAL),
				
				// NEW: Add Tenant ID field (Optional)
				(new DefinitionParameter(
					'tenant',
					$l10n->t('Tenant ID (Optional)')
				))
					->setFlag(DefinitionParameter::FLAG_OPTIONAL), // User can leave it blank				
            ]);
    }
}

