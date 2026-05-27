<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Listener;

use OCA\Files_External\Event\StorageCreatedEvent;
use OCA\Files_External\Event\StorageUpdatedEvent;
use OCA\Sharepoint2\BackgroundJob\WarmupCacheJob;
use OCP\BackgroundJob\IJobList;
use OCP\EventDispatcher\Event;
use OCP\EventDispatcher\IEventListener;

/** @template-implements IEventListener<Event|StorageCreatedEvent|StorageUpdatedEvent> */
class StorageChangedListener implements IEventListener {
    public function __construct(
        private IJobList $jobList,
    ) {
    }

    public function handle(Event $event): void {
        if ($event instanceof StorageCreatedEvent) {
            $storage = $event->getNewConfig();
        } elseif ($event instanceof StorageUpdatedEvent) {
            $storage = $event->getNewConfig();
        } else {
            return;
        }

        if ($storage->getBackend()->getIdentifier() !== 'sharepoint2') {
            return;
        }

        $argument = [
            'mount_id' => (int)$storage->getId(),
            'full' => false,
            'path' => '',
        ];

        if (!$this->jobList->has(WarmupCacheJob::class, $argument)) {
            $this->jobList->add(WarmupCacheJob::class, $argument);
        }
    }
}
