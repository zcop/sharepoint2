/* global OCA, OC */

$(document).ready(function () {
	var backendId = 'sharepoint2';
	var backendUrl = OC.generateUrl('apps/' + backendId + '/oauth');

	function displayGranted($tr) {
		$tr.find('.configuration input.auth-param')
			.attr('disabled', 'disabled')
			.addClass('disabled-success');
	}

	if (!OCA.Files_External ||
		!OCA.Files_External.Settings ||
		!OCA.Files_External.Settings.mountConfig ||
		!OCA.Files_External.Settings.mountConfig.whenSelectAuthMechanism) {
		return;
	}

	/**
	 * Hook into auth mechanism selection for external storage rows.
	 */
	OCA.Files_External.Settings.mountConfig.whenSelectAuthMechanism(function ($tr, authMechanism, scheme, onCompletion) {
		if (authMechanism === 'oauth2::oauth2' && $tr.hasClass(backendId)) {
			var config = $tr.find('.configuration');

			setTimeout(function () {
				config.find('[name="oauth2_grant"]')
					.attr('name', 'oauth2_grant_sharepoint2');

				var $lib = config.find('[data-parameter="library_path"]');
				if ($lib.length) {
					if ($lib.val().trim() === '') {
						$lib.val('Documents');
					}
					$lib.prop('readonly', true);
				}
			}, 50);

			if (onCompletion && typeof onCompletion.then === 'function') {
				onCompletion.then(function () {
					var configured = $tr.find('[data-parameter="configured"]');

					if ($(configured).val() === 'true') {
						if (localStorage.getItem('sharepoint2_oauth2')) {
							localStorage.removeItem('sharepoint2_oauth2');
						}
						displayGranted($tr);
					} else {
						var client_id = $tr.find('.configuration [data-parameter="client_id"]').val().trim();
						var client_secret = $tr.find('.configuration [data-parameter="client_secret"]').val().trim();
						var tenant = $tr.find('.configuration [data-parameter="tenant"]').val().trim();
						
						if (localStorage.getItem('sharepoint2_oauth2')) {
							client_secret = atob(localStorage.getItem('sharepoint2_oauth2'));
						}

						var params = {};
						window.location.href.replace(/[?&]+([^=&]+)=([^&]*)/gi, function (m, key, value) {
							params[key] = decodeURIComponent(value);
						});

						if (
							params.code !== undefined &&
							typeof client_id === 'string' &&
							client_id !== '' &&
							typeof client_secret === 'string' &&
							client_secret !== ''
						) {
							$('.configuration').trigger('sharepoint2_oauth_step2', [{
								backend_id: backendId,
								client_id: client_id,
								client_secret: client_secret,
								tenant: tenant, // --- FIX 2: Pass it to the event ---								
								redirect: location.protocol + '//' + location.host + location.pathname,
								tr: $tr,
								code: params.code || '',
								state: params.state || ''
							}]);
						}
					}
				});
			}
		}
	});

	/**
	 * Click handler for our custom Grant button.
	 */
	$('#externalStorage').on('click', '[name="oauth2_grant_sharepoint2"]', function (event) {
		event.preventDefault();
		var tr = $(this).closest('tr');
		var client_id = tr.find('.configuration [data-parameter="client_id"]').val().trim();
		var client_secret = tr.find('.configuration [data-parameter="client_secret"]').val().trim();
		var tenant = tr.find('.configuration [data-parameter="tenant"]').val().trim();
		
		if (client_id !== '' && client_secret !== '') {
			$('.configuration').trigger('sharepoint2_oauth_step1', [{
				backend_id: backendId,
				client_id: client_id,
				client_secret: client_secret,
				tenant: tenant,
				redirect: location.protocol + '//' + location.host + location.pathname,
				tr: tr
			}]);
		}
	});

	/**
	 * STEP 1: ask backend for auth URL and redirect browser to Microsoft.
	 */
	$('.configuration').on('sharepoint2_oauth_step1', function (event, data) {
		if (data['backend_id'] !== backendId) {
			return false;
		}
		OCA.Files_External.Settings.OAuth2.getSharePointAuthUrl(backendUrl, data);
	});

	/**
	 * STEP 2: after redirect back, verify the code via backend and save token.
	 */
	$('.configuration').on('sharepoint2_oauth_step2', function (event, data) {
		if (data['backend_id'] !== backendId || data['code'] === undefined) {
			return false;
		}
		OCA.Files_External.Settings.OAuth2.sharePointVerifyCode(backendUrl, data)
			.fail(function (message) {
				console.log('SharePoint OAuth2 verify failed: ' + message);
				OC.dialogs.alert(
					message,
					t('sharepoint2', 'Error verifying OAuth2 Code for SharePoint Online')
				);
			});
	});

	OCA.Files_External.Settings.OAuth2 = OCA.Files_External.Settings.OAuth2 || {};

	/**
	 * Helper: STEP 1 – backend builds Microsoft auth URL, we redirect browser.
	 */
	OCA.Files_External.Settings.OAuth2.getSharePointAuthUrl = function (backendUrl, data) {
		var $tr = data['tr'];
		var configured = $tr.find('[data-parameter="configured"]');
		var token = $tr.find('.configuration [data-parameter="token"]');
		var client_secret = data['client_secret'];

		if (localStorage.getItem('sharepoint2_oauth2')) {
			client_secret = atob(localStorage.getItem('sharepoint2_oauth2'));
		}

		$.post(backendUrl, {
			step: 1,
			client_id: data['client_id'],
			client_secret: client_secret,
			tenant: data['tenant'],
			redirect: data['redirect']
		})
			.done(function (result) {
				if (result && result.status === 'success') {
					$(configured).val('false');
					$(token).val('false');

					OCA.Files_External.Settings.mountConfig.saveStorageConfig($tr, function () {
						if (!result.data.url) {
							OC.dialogs.alert(
								'Auth URL not set',
								t('files_external', 'Error getting OAuth2 URL for ' + data['backend_id'])
							);
						} else {
							localStorage.setItem('sharepoint2_oauth2', btoa(data['client_secret']));
							window.location = result.data.url;
						}
					});
				} else {
					OC.dialogs.alert(
						result && result.data && result.data.message ? result.data.message : 'Unknown error',
						t('files_external', 'Error getting OAuth2 URL for ' + data['backend_id'])
					);
				}
			})
			.fail(function (xhr, status, error) {
				console.log('Error during OAuth2 get URL', status, error);
			});
	};

	/**
	 * Helper: STEP 2 – send code to backend, get token, mark mount as configured.
	 * Returns a jQuery Deferred.
	 */
	OCA.Files_External.Settings.OAuth2.sharePointVerifyCode = function (backendUrl, data) {
		var deferredObject = $.Deferred();

		var $tr = data['tr'];
		var configured = $tr.find('[data-parameter="configured"]');
		var token = $tr.find('.configuration [data-parameter="token"]');
		var client_secret = data['client_secret'];

		if (localStorage.getItem('sharepoint2_oauth2')) {
			client_secret = atob(localStorage.getItem('sharepoint2_oauth2'));
		}

		$.post(backendUrl, {
			step: 2,
			client_id: data['client_id'],
			client_secret: client_secret,
			tenant: data['tenant'], 
			redirect: data['redirect'],
			code: data['code'],
			state: data['state']
		})
			.done(function (result) {
				if (result && result.status === 'success') {
					$(token).val(result.data.token);
					$(configured).val('true');

					OCA.Files_External.Settings.mountConfig.saveStorageConfig($tr, function (status) {
						if (status) {
							displayGranted($tr);
						}
						deferredObject.resolve(status);
					});
				} else {
					var msg = result && result.data && result.data.message ? result.data.message : 'Unknown error';
					deferredObject.reject(msg);
				}
			})
			.fail(function (xhr, status, error) {
				console.log('Error during OAuth2 verify code', status, error);
				deferredObject.reject(error || status);
			});

		return deferredObject.promise();
	};
});
