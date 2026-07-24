import { build, context, type BuildOptions } from 'esbuild';

const isProd = process.argv.includes('--mode=production');

const options: BuildOptions = {
    entryPoints: ['./src/extension.ts'],
    bundle: true,
    outfile: 'out/extension.js',
    external: ['vscode'],
    format: 'cjs',
    platform: 'node',
    target: ['node20'],
    minify: isProd,
    sourcemap: !isProd,
    logOverride: {
        'duplicate-object-key': 'silent',
        'suspicious-boolean-not': 'silent',
    },
    plugins: [{
        name: 'build-notice',
        setup(esbuild) {
            esbuild.onStart(() => {
                console.log('desktop extension build start');
            });
            esbuild.onEnd(result => {
                if (result.errors.length === 0) {
                    console.log('desktop extension build success');
                }
            });
        },
    }],
};

async function buildDesktopExtension(): Promise<void> {
    if (isProd) {
        await build(options);
        return;
    }
    const buildContext = await context(options);
    await buildContext.watch();
}

void buildDesktopExtension().catch(error => {
    console.error(error);
    process.exitCode = 1;
});
