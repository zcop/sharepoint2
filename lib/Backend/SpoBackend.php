<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Backend;

use OCA\Files_External\Lib\Backend\Backend;
use OCA\Files_External\Lib\Auth\AuthMechanism;
use OCA\Files_External\Lib\DefinitionParameter;
use OCP\IL10N;

/**
 * SharePoint Online (OAuth2) backend definition.
 *
 * For now:
 *  - Storage class is still Local (placeholder)
 *  - Library path is fixed to "Documents" (GUI shows textbox but readonly)
 */
class SpoBackend extends Backend {

    public function __construct() {
        $appWebPath = \OC_App::getAppWebPath('sharepoint2');

	/** @var IL10N $l */
        $l = \OC::$server->getL10N('sharepoint2');

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
                    $l->t('SharePoint Site URL (e.g. https://tenant.sharepoint.com/sites/MySite)')
                ),

                // Library path, e.g. "Documents" or "Documents/SubFolder"
                (new DefinitionParameter(
                    'library',
                    $l->t('Library path (e.g. Documents or Documents/SubFolder)')
                ))
                    ->setFlag(DefinitionParameter::FLAG_OPTIONAL),
            ]);
    }
}

