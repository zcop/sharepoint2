<?php
declare(strict_types=1);

return [
    'routes' => [
        // POST /apps/sharepoint2/oauth
        [
            'name' => 'oauth#receiveToken',
            'url'  => '/oauth',
            'verb' => 'POST',
        ],
    ],
];
