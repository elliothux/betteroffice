import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import {
  NODE_BINDINGS,
  NODE_BINDING_NAMES,
  bindingVersion,
  pendingPublishNames,
  platformPackageVersions
} from './node-bindings.mjs';

const releaseWorkflow = fileURLToPath(new URL('../.github/workflows/release.yml', import.meta.url));
const publishWorkflow = fileURLToPath(
  new URL('../.github/workflows/publish-node-binding.yml', import.meta.url)
);
const distWorkflow = fileURLToPath(new URL('../.github/workflows/node-dist.yml', import.meta.url));

describe('Node binding registry', () => {
  test('registers every native format and platform package', () => {
    expect(NODE_BINDING_NAMES).toEqual(['docx', 'pptx', 'xlsx']);
    expect(NODE_BINDINGS).toEqual(NODE_BINDING_NAMES.map((name) => `bindings/node-${name}`));
    expect(platformPackageVersions()).toHaveLength(NODE_BINDINGS.length * 5);
    const versions = new Set(NODE_BINDINGS.map(bindingVersion));
    expect(versions.size).toBe(1);
    expect(new Set(platformPackageVersions().map((entry) => entry.version))).toEqual(versions);
  });

  test('detects only unpublished root versions', async () => {
    const pending = await pendingPublishNames({
      fetchImpl: async (url: string) => {
        const name = decodeURIComponent(new URL(url).pathname.slice(1));
        return name.includes('pptx')
          ? new Response('{"error":"missing"}', { status: 404 })
          : Response.json({ versions: { '0.0.1': {} } });
      }
    });
    expect(pending).toEqual(['pptx']);
  });
});

describe('Node binding release wiring', () => {
  const release = Bun.YAML.parse(readFileSync(releaseWorkflow, 'utf8')) as any;
  const publish = Bun.YAML.parse(readFileSync(publishWorkflow, 'utf8')) as any;
  const dist = Bun.YAML.parse(readFileSync(distWorkflow, 'utf8')) as any;

  test('release dispatches and waits for the dedicated publisher', () => {
    const step = release.jobs.release.steps.find(
      (value: any) => value.name === 'Publish Node native bindings'
    );
    expect(step.run).toContain('scripts/node-bindings.mjs --pending');
    expect(step.run).toContain('publish-node-binding.yml/dispatches');
    expect(step.run).toContain('gh run watch');
  });

  test('publisher uses OIDC and builds all declared platforms at the requested commit', () => {
    expect(publish.jobs.publish.permissions['id-token']).toBe('write');
    expect(publish.jobs.publish.environment).toBe('npm-${{ inputs.binding }}');
    expect(publish.jobs.publish.steps.some((step: any) => step.run?.includes('NPM_TOKEN'))).toBe(
      true
    );
    expect(dist.jobs.bindings.strategy.matrix.platform).toHaveLength(5);
    expect(dist.jobs.bindings.steps[0].with.ref).toBe('${{ inputs.sha }}');
  });
});
