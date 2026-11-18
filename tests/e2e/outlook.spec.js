import { test, expect } from '@playwright/test';
import { createServer } from 'node:http';
import { startApp, closeAndCleanup, invokeHandler } from './helpers/electronApp.js';

// Exercise the real application with deterministic pages. No Microsoft account
// or tenant data is needed; live Microsoft sign-in is a separate manual check.
test.describe('Outlook workspace tab', () => {
  test.describe.configure({ mode: 'serial' });
  let ctx;
  let server;
  let outlookRequests = 0;
  test.beforeAll(async () => {
    server = createServer((_req, res) => {
      res.setHeader('Content-Type', 'text/html');
      res.end('<!doctype html><title>Teams fixture</title><input id="draft" aria-label="Teams draft"><script>window.bootId = Math.random(); window.ticks = 0; setInterval(() => window.ticks++, 100);</script>');
    });
    await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
    ctx = await startApp({
      prefix: 'teams-outlook-e2e-', allowEval: true,
      config: {
        app: { url: `http://127.0.0.1:${server.address().port}/` },
        outlook: { enabled: true },
      },
    });
    await expect.poll(() => ctx.electronApp.evaluate(({ BaseWindow }) =>
      BaseWindow.getAllWindows().flatMap((w) => w.contentView.children.map((v) => v.webContents?.getURL()))
    )).toContainEqual(expect.stringContaining('/outlook/tabs.html'));
    await ctx.electronApp.evaluate(({ BaseWindow }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      globalThis.outlookRequestCount = 0;
      win.webContents.session.protocol.handle('https', (request) => {
        globalThis.outlookRequestCount++;
        if (new URL(request.url).pathname === '/offline') return Response.error();
        return new Response('<!doctype html><title>Outlook fixture</title><input id="draft" aria-label="Mail draft"><a id="auth" target="_blank" href="https://login.microsoftonline.com/common/oauth2">Sign in</a>', { headers: { 'content-type': 'text/html' } });
      });
    });
  });
  test.afterAll(async () => {
    if (ctx) await closeAndCleanup(ctx);
    if (server) await new Promise((resolve) => server.close(resolve));
  });

  async function probe(expression, target = 'bar') {
    return ctx.electronApp.evaluate(async ({ BaseWindow }, { expression, target }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const view = win.contentView.children.find((v) => {
        const url = v.webContents?.getURL();
        return target === 'bar' ? url?.endsWith('/outlook/tabs.html') : url?.startsWith('https://outlook.office.com');
      });
      const contents = target === 'teams' ? win.webContents : view?.webContents;
      return contents ? contents.executeJavaScript(expression) : null;
    }, { expression, target });
  }

  async function select(tab) {
    // Deliver mouse input to the actual isolated tab-strip renderer.
    await ctx.electronApp.evaluate(async ({ BaseWindow }, tab) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const bar = win.contentView.children.find((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html'));
      const point = await bar.webContents.executeJavaScript(`(() => { const r = document.getElementById(${JSON.stringify(tab)}).getBoundingClientRect(); return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) }; })()`);
      bar.webContents.sendInputEvent({ type: 'mouseDown', button: 'left', clickCount: 1, ...point });
      bar.webContents.sendInputEvent({ type: 'mouseUp', button: 'left', clickCount: 1, ...point });
    }, tab);
  }

  test('loads Outlook lazily, shares the session, and isolates remote code', async () => {
    expect(await ctx.electronApp.evaluate(() => globalThis.outlookRequestCount)).toBe(0);
    await select('outlook');
    await expect.poll(() => probe('document.getElementById("outlook").getAttribute("aria-selected")')).toBe('true');
    await expect.poll(() => probe('document.title', 'outlook')).toBe('Outlook fixture');
    const prefs = await ctx.electronApp.evaluate(({ BaseWindow }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const mail = win.contentView.children.find((v) => v.webContents?.getURL().startsWith('https://outlook.office.com'));
      return { sameSession: mail.webContents.session === win.webContents.session, prefs: mail.webContents.getLastWebPreferences() };
    });
    expect(prefs.sameSession).toBe(true);
    expect(prefs.prefs.contextIsolation).toBe(true);
    expect(prefs.prefs.sandbox).toBe(true);
    expect(prefs.prefs.nodeIntegration).toBe(false);
    expect(prefs.prefs.preload || '').toBe('');
    expect(await probe('typeof window.electronAPI', 'outlook')).toBe('undefined');
    outlookRequests = await ctx.electronApp.evaluate(() => globalThis.outlookRequestCount);
  });

  test('switches repeatedly without losing drafts or reloading Teams', async () => {
    const original = await probe('document.getElementById("draft").value = "Teams draft"; ({ boot: window.bootId, ticks: window.ticks })', 'teams');
    await probe('document.getElementById("draft").value = "Mail draft"', 'outlook');
    for (let i = 0; i < 3; i++) {
      await select('teams');
      await expect.poll(() => probe('document.getElementById("teams").getAttribute("aria-selected")')).toBe('true');
      await select('outlook');
      await expect.poll(() => probe('document.getElementById("outlook").getAttribute("aria-selected")')).toBe('true');
    }
    expect(await probe('window.bootId', 'teams')).toBe(original.boot);
    await expect.poll(() => probe('window.ticks', 'teams')).toBeGreaterThan(original.ticks);
    expect(await probe('document.getElementById("draft").value', 'teams')).toBe('Teams draft');
    expect(await probe('document.getElementById("draft").value', 'outlook')).toBe('Mail draft');
    expect(await ctx.electronApp.evaluate(() => globalThis.outlookRequestCount)).toBe(outlookRequests);
  });

  test('reserves a full-width header above both apps when resized', async () => {
    const layout = await ctx.electronApp.evaluate(({ BaseWindow }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      win.setContentSize(900, 650);
      return new Promise((resolve) => setTimeout(async () => {
        const [width, height] = win.getContentSize();
        const views = win.contentView.children;
        resolve({ width, height, bar: views.at(-1).getBounds(), teams: win.teamsView.getBounds(), viewport: await win.webContents.executeJavaScript('({ width: innerWidth, height: innerHeight })'), mail: views.find((v) => v.webContents?.getURL().startsWith('https://outlook.office.com')).getBounds() });
      }, 200));
    });
    expect(layout.bar).toEqual({ x: 0, y: 0, width: layout.width, height: 52 });
    expect(layout.mail).toEqual({ x: 0, y: 52, width: layout.width, height: layout.height - 52 });
    expect(layout.teams).toEqual(layout.mail);
    expect(layout.viewport).toEqual({ width: layout.width, height: layout.height - 52 });
  });

  test('shows load errors and retries successfully', async () => {
    await ctx.electronApp.evaluate(async ({ BaseWindow }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const mail = win.contentView.children.find((v) => v.webContents?.getURL().startsWith('https://outlook.office.com'));
      await mail.webContents.loadURL('https://outlook.office.com/offline').catch(() => {});
    });
    await expect.poll(() => probe('document.getElementById("status").textContent')).toBe('Load failed');
    await select('reload');
    await expect.poll(() => probe('document.title', 'outlook')).toBe('Outlook fixture');
    await expect.poll(() => probe('document.getElementById("status").textContent')).toBe('');
  });

  test('supports sandboxed Microsoft authentication popups', async () => {
    const before = await ctx.electronApp.evaluate(({ BaseWindow }) => BaseWindow.getAllWindows().length);
    await probe('document.getElementById("auth").click()', 'outlook');
    await expect.poll(() => ctx.electronApp.evaluate(({ BaseWindow }) => BaseWindow.getAllWindows().length)).toBe(before + 1);
    const readPrefs = () => ctx.electronApp.evaluate(({ BaseWindow }) => {
      const popup = BaseWindow.getAllWindows().find((w) => w.webContents.getURL().startsWith('https://login.microsoftonline.com'));
      return popup?.webContents.getLastWebPreferences();
    });
    await expect.poll(readPrefs).toMatchObject({ sandbox: true, nodeIntegration: false });
    const prefs = await readPrefs();
    expect(prefs.sandbox).toBe(true);
    expect(prefs.nodeIntegration).toBe(false);
    expect(prefs.preload || '').toBe('');
  });

  test('rejects workspace commands from another renderer', async () => {
    await ctx.electronApp.evaluate(({ BaseWindow, ipcMain }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const bar = win.contentView.children.find((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html'));
      // Even a local URL is insufficient: the sender must be this tab strip.
      ipcMain.emit('workspace-tab-action', { sender: win.webContents, senderFrame: bar.webContents.mainFrame }, { tab: 'teams' });
    });
    expect(await probe('document.getElementById("outlook").getAttribute("aria-selected")')).toBe('true');
  });

  test('supports keyboard tab navigation', async () => {
    await ctx.electronApp.evaluate(async ({ BaseWindow }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const bar = win.contentView.children.find((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html'));
      await bar.webContents.executeJavaScript('document.getElementById("outlook").focus()');
      bar.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Left' });
      bar.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Left' });
    });
    await expect.poll(() => probe('document.getElementById("teams").getAttribute("aria-selected")')).toBe('true');
  });
});

test('Outlook coexists with the profile switcher and returns to the chosen Teams profile', async () => {
  const ctx = await startApp({
    prefix: 'teams-outlook-profiles-', allowEval: true,
    config: { outlook: { enabled: true }, multiAccount: { enabled: true } },
  });
  try {
    const added = await invokeHandler(ctx.electronApp, 'profile-add', [{ name: 'Outlook test profile' }]);
    expect(added.result.id).toBeTruthy();
    const initial = await ctx.electronApp.evaluate(async ({ BaseWindow }) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const bar = win.contentView.children.find((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html'));
      const x = bar.getBounds().x;
      await bar.webContents.executeJavaScript('window.workspaceTabs.select("outlook")');
      return x;
    });
    expect(initial).toBe(0);
    await invokeHandler(ctx.electronApp, 'profile-switch', [added.result.id]);
    await expect.poll(() => ctx.electronApp.evaluate(async ({ BaseWindow, session }, partition) => {
      const win = BaseWindow.getAllWindows().find((w) => w.contentView.children.some((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html')));
      const bar = win.contentView.children.find((v) => v.webContents?.getURL().endsWith('/outlook/tabs.html'));
      const profile = win.contentView.children.find((v) => v !== win.teamsView && v.getVisible() && v.webContents?.session === session.fromPartition(partition));
      return {
        selected: await bar.webContents.executeJavaScript('document.getElementById("teams").getAttribute("aria-selected")'),
        profileY: profile?.getBounds().y,
      };
    }, added.result.partition)).toEqual({ selected: 'true', profileY: 52 });
  } finally {
    await closeAndCleanup(ctx);
  }
});
