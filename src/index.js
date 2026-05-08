import { getRenderer as ean13Renderer } from './ean13.js';
import { getRenderer as code128Renderer } from './code128.js';
import { getRenderer as qrRenderer } from './qrcode.js';
import { encodePNG } from './png.js';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

function jsonError(status, message) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
  });
}

function selectRenderer(type, value, width, height) {
  switch (type.toUpperCase().replace(/[-_ ]/g, '')) {
    case 'EAN13':    return ean13Renderer(value, width, height);
    case 'CODE128':
    case 'CODE128B':
    case 'C128':     return code128Renderer(value, width, height);
    case 'QR':
    case 'QRCODE':   return qrRenderer(value, width, height);
    default: throw new Error(`Unknown type "${type}". Supported: EAN13, CODE128, QR`);
  }
}

// Serve a manifest.xml with the current worker's origin baked in
function serveManifest(origin) {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>A1B2C3D4-E5F6-7890-ABCD-EF1234567890</Id>
  <Version>1.0.0</Version>
  <ProviderName>Excel Barcode</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Barcode Inserter"/>
  <Description DefaultValue="Insert EAN-13, CODE128 and QR barcodes embedded directly into your workbook"/>
  <SupportUrl DefaultValue="${origin}"/>
  <Hosts><Host Name="Workbook"/></Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${origin}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="BarcodeGroup">
                <Label resid="CommandsGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="OpenPaneButton">
                  <Label resid="OpenPaneButton.Label"/>
                  <Supertip>
                    <Title resid="OpenPaneButton.Label"/>
                    <Description resid="OpenPaneButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>BarcodePane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="${origin}/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="${origin}/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="${origin}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url"  DefaultValue="${origin}/commands.html"/>
        <bt:Url id="Taskpane.Url"  DefaultValue="${origin}/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="CommandsGroup.Label"    DefaultValue="Barcodes"/>
        <bt:String id="OpenPaneButton.Label"   DefaultValue="Insert Barcode"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="OpenPaneButton.Tooltip" DefaultValue="Open the Barcode Inserter panel"/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;
  return new Response(xml, {
    headers: { ...CORS, 'Content-Type': 'application/xml; charset=utf-8' },
  });
}

// Minimal placeholder icon (1x1 blue PNG, base64)
const ICON_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
function iconBytes() {
  const b = atob(ICON_B64);
  const arr = new Uint8Array(b.length);
  for (let i = 0; i < b.length; i++) arr[i] = b.charCodeAt(i);
  return arr;
}

export default {
  async fetch(request, _env, ctx) {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS });
    }

    // Barcode API
    if (url.pathname === '/barcode') {
      const value  = url.searchParams.get('value');
      const type   = url.searchParams.get('type') || 'EAN13';
      const width  = Math.min(2000, Math.max(1, parseInt(url.searchParams.get('width')  || '300', 10)));
      const height = Math.min(1000, Math.max(1, parseInt(url.searchParams.get('height') || '150', 10)));

      if (!value) return jsonError(400, 'Missing required parameter: value');

      let renderer;
      try {
        renderer = selectRenderer(type, value, width, height);
      } catch (e) {
        return jsonError(400, e.message);
      }

      const png = encodePNG(renderer.width, renderer.height, renderer.isDark);

      return new Response(png, {
        headers: {
          ...CORS,
          'Content-Type': 'image/png',
          'Cache-Control': 'public, max-age=86400',
          'X-Barcode-Width':  String(renderer.width),
          'X-Barcode-Height': String(renderer.height),
        },
      });
    }

    // Dynamic manifest (origin baked in automatically)
    if (url.pathname === '/manifest.xml') {
      return serveManifest(url.origin);
    }

    // Placeholder icons
    if (url.pathname.startsWith('/assets/icon-')) {
      return new Response(iconBytes(), {
        headers: { 'Content-Type': 'image/png', 'Cache-Control': 'public, max-age=86400' },
      });
    }

    // Root and /guide — redirect to installation guide
    if (url.pathname === '/' || url.pathname === '/guide') {
      return Response.redirect(`${url.origin}/guide.html`, 302);
    }
    if (url.pathname === '/guide-el') {
      return Response.redirect(`${url.origin}/guide-el.html`, 302);
    }

    // Fall through to static assets (taskpane.html, taskpane.js, commands.html)
    return _env.ASSETS.fetch(request);
  },
};
