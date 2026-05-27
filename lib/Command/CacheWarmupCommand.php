<?php
declare(strict_types=1);

namespace OCA\Sharepoint2\Command;

use OC\Core\Command\Base;
use OCA\Sharepoint2\Service\CacheWarmupService;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;

class CacheWarmupCommand extends Base {
    public function __construct(
        private CacheWarmupService $cacheWarmupService,
    ) {
        parent::__construct();
    }

    protected function configure(): void {
        $this
            ->setName('sharepoint2:cache:warmup')
            ->setDescription('Warm up SharePoint2 external storage cache (root + level-1 split scan)')
            ->addArgument('mount_id', InputArgument::REQUIRED, 'SharePoint2 mount id from files_external:list')
            ->addOption('full', null, InputOption::VALUE_NONE, 'Run full scan for this mount')
            ->addOption('path', null, InputOption::VALUE_OPTIONAL, 'Scan a specific path only', '');

        parent::configure();
    }

    protected function execute(InputInterface $input, OutputInterface $output): int {
        $mountId = (int)$input->getArgument('mount_id');
        $full = (bool)$input->getOption('full');
        $path = trim((string)$input->getOption('path'));

        try {
            $result = $this->cacheWarmupService->warmupMountById($mountId, $full, $path);
        } catch (\Throwable $e) {
            $output->writeln('<error>' . $e->getMessage() . '</error>');
            return 1;
        }

        $output->writeln('<info>Mode: ' . $result['mode'] . '</info>');
        $output->writeln('<info>Mount: ' . $result['mount_id'] . '</info>');

        if ($result['path'] !== '') {
            $output->writeln('<info>Path scanned: ' . $result['path'] . '</info>');
        }

        if (!$result['full_scan'] && $result['mode'] === 'warmup') {
            $output->writeln('<info>Level-1 directories discovered: ' . $result['level1_dirs'] . '</info>');
            $output->writeln('<info>Level-1 directories scanned: ' . $result['scanned_dirs'] . '</info>');
        }

        return 0;
    }
}
