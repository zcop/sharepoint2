<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Auth;

use OCA\Files_External\Lib\Auth\AuthMechanism;
use OCP\IL10N;

class OAuth2Mechanism extends AuthMechanism {
	public function __construct(IL10N $l10n) {
		$this
			->setIdentifier('oauth2::oauth2')
			->setScheme(self::SCHEME_OAUTH2)
			->setText($l10n->t('OAuth2'));
	}
}

