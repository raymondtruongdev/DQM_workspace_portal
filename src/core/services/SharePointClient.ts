import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISharePointFileInfo {
  name: string;
  serverRelativeUrl: string;
  timeLastModified: string;
  length: number;
  eTag?: string;
}

/**
 * Generic SharePoint client (REST via SPHttpClient v1).
 * Module-specific services (ProgressService, UserProfileService...) are built on top of it.
 */
export class SharePointClient {
  private _spHttpClient: SPHttpClient;

  constructor(spHttpClient: SPHttpClient) {
    this._spHttpClient = spHttpClient;
  }

  private _escapeUrl(url: string): string {
    return url.replace(/'/g, "''");
  }

  private _itemsUrl(siteUrl: string, listTitle: string, id?: number): string {
    const base = `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`;
    return id === undefined ? base : `${base}(${id})`;
  }

  /**
   * Fetch items from a list
   */
  public async getItems<T>(siteUrl: string, listTitle: string, top = 5000): Promise<T[]> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      `${this._itemsUrl(siteUrl, listTitle)}?$top=${top}`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`SharePoint API Error (${response.status}): ${errText}`);
    }

    const json = await response.json();
    return (json.value || []) as T[];
  }

  /**
   * Query items with OData options, e.g. "$filter=Unit eq 'A'&$orderby=Id desc&$top=1".
   */
  public async queryItems<T>(siteUrl: string, listTitle: string, query: string): Promise<T[]> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      `${this._itemsUrl(siteUrl, listTitle)}?${query}`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`SharePoint API Error (${response.status}): ${errText}`);
    }

    const json = await response.json();
    return (json.value || []) as T[];
  }

  /**
   * Visible columns of a list as { displayTitle: internalName } (internal names can differ from titles, e.g. "Unit0").
   */
  public async getListFieldMap(siteUrl: string, listTitle: string): Promise<Record<string, string>> {
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${listTitle}')/fields?$filter=Hidden eq false&$select=Title,InternalName`,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );
    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`SharePoint API Error (${response.status}): ${errText}`);
    }
    const json = await response.json();
    const map: Record<string, string> = {};
    ((json.value || []) as { Title: string; InternalName: string }[]).forEach((f) => {
      if (/^LinkTitle/i.test(f.InternalName)) return; // cột hiển thị (có menu) của Title, không ghi được
      map[f.Title.toLowerCase()] = f.InternalName;
    });
    return map;
  }

  /**
   * Create an item in a list
   */
  public async createItem<T>(siteUrl: string, listTitle: string, data: Partial<T>): Promise<void> {
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      this._itemsUrl(siteUrl, listTitle),
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(data)
      }
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to create item (${response.status}): ${errText}`);
    }
  }

  /**
   * Update (MERGE) an item in a list
   */
  public async updateItem<T>(siteUrl: string, listTitle: string, id: number, data: Partial<T>): Promise<void> {
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      this._itemsUrl(siteUrl, listTitle, id),
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'X-HTTP-Method': 'MERGE',
          'If-Match': '*'
        },
        body: JSON.stringify(data)
      }
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to update item: ${errText}`);
    }
  }

  /**
   * Fetch binary contents and ETag of a file. Returns undefined if not found (404).
   */
  public async getFileWithMeta(
    siteUrl: string,
    serverRelativeUrl: string
  ): Promise<{ buffer: ArrayBuffer; eTag: string } | undefined> {
    const safeUrl = this._escapeUrl(serverRelativeUrl);
    const endpoint = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${safeUrl}')/$value`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache'
        }
      }
    );

    if (response.status === 404) {
      return undefined;
    }
    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to get file (${response.status}): ${errText}`);
    }

    const buffer = await response.arrayBuffer();
    const eTag = response.headers.get('ETag') || response.headers.get('etag') || '';
    return { buffer, eTag };
  }

  /**
   * Fetch binary contents of a file as ArrayBuffer.
   */
  public async getFileBuffer(siteUrl: string, serverRelativeUrl: string): Promise<ArrayBuffer | undefined> {
    const res = await this.getFileWithMeta(siteUrl, serverRelativeUrl);
    return res ? res.buffer : undefined;
  }

  /**
   * Upload binary data to a SharePoint folder.
   */
  public async uploadFile(
    siteUrl: string,
    folderServerRelativeUrl: string,
    fileName: string,
    buffer: ArrayBuffer,
    overwrite = true,
    eTag?: string
  ): Promise<{ serverRelativeUrl: string; eTag?: string }> {
    const safeFolder = this._escapeUrl(folderServerRelativeUrl);
    const safeFileName = this._escapeUrl(fileName);
    const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${safeFolder}')/Files/add(url='${safeFileName}',overwrite=${overwrite})`;

    const headers: Record<string, string> = {
      'Accept': 'application/json;odata=nometadata',
      'Content-Type': 'application/octet-stream'
    };
    if (eTag) {
      headers['If-Match'] = eTag;
    }

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers,
        body: buffer
      }
    );

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to upload file (${response.status}): ${errText}`);
    }

    const json = await response.json();
    return {
      serverRelativeUrl: json.ServerRelativeUrl || `${folderServerRelativeUrl}/${fileName}`,
      eTag: json.ETag || response.headers.get('ETag') || undefined
    };
  }

  /**
   * Delete a file at the given server-relative URL.
   */
  public async deleteFile(siteUrl: string, serverRelativeUrl: string): Promise<void> {
    const safeUrl = this._escapeUrl(serverRelativeUrl);
    const endpoint = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${safeUrl}')`;
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'X-HTTP-Method': 'DELETE',
          'If-Match': '*'
        }
      }
    );

    if (!response.ok && response.status !== 404) {
      const errText = await response.text();
      throw new Error(`Failed to delete file (${response.status}): ${errText}`);
    }
  }

  /**
   * Check if a folder exists at the given server-relative URL.
   */
  public async folderExists(siteUrl: string, folderServerRelativeUrl: string): Promise<boolean> {
    try {
      const safeUrl = this._escapeUrl(folderServerRelativeUrl);
      const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${safeUrl}')?$select=Exists`;
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );
      if (!response.ok) return false;
      const data = await response.json();
      if (data && typeof data.Exists === 'boolean') {
        return data.Exists;
      }
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Ensure a folder hierarchy exists by creating missing folders step by step.
   */
  public async ensureFolder(siteUrl: string, folderServerRelativeUrl: string): Promise<void> {
    // Normalise slashes
    const cleanPath = folderServerRelativeUrl.replace(/\\/g, '/').replace(/\/+/g, '/').replace(/\/$/, '');
    if (!cleanPath) return;

    // Check if whole folder already exists
    if (await this.folderExists(siteUrl, cleanPath)) {
      return;
    }

    // Determine the web's server-relative path (e.g. /sites/metro-checkin-dev)
    let webServerRel = '';
    try {
      if (siteUrl.startsWith('http://') || siteUrl.startsWith('https://')) {
        webServerRel = new URL(siteUrl).pathname.replace(/\/+$/, '').toLowerCase();
      } else {
        webServerRel = '/' + siteUrl.replace(/^\/+/, '').replace(/\/+$/, '').toLowerCase();
      }
    } catch {
      webServerRel = '';
    }

    const parts = cleanPath.split('/').filter(Boolean);
    let currentPath = '';

    for (let i = 0; i < parts.length; i++) {
      currentPath += '/' + parts[i];

      // Skip root /sites or the site web itself
      const currentLower = currentPath.toLowerCase();
      if (currentLower === '/sites') {
        continue;
      }
      if (webServerRel && (currentLower === webServerRel || webServerRel.startsWith(currentLower + '/'))) {
        continue;
      }

      // Check if currentPath already exists
      const alreadyExists = await this.folderExists(siteUrl, currentPath);
      if (alreadyExists) {
        continue;
      }

      // Create folder at currentPath
      let created = false;
      let lastError = '';

      // Method 1: POST /_api/web/folders with verbose metadata
      try {
        const res1 = await this._spHttpClient.post(
          `${siteUrl}/_api/web/folders`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose'
            },
            body: JSON.stringify({
              '__metadata': { 'type': 'SP.Folder' },
              'ServerRelativeUrl': currentPath
            })
          }
        );
        if (res1.ok) {
          created = true;
        } else {
          lastError = await res1.text();
        }
      } catch (err: unknown) {
        lastError = err instanceof Error ? err.message : String(err);
      }

      // Method 2: POST /_api/web/folders/addUsingPath
      if (!created) {
        try {
          const safeCurr = this._escapeUrl(currentPath);
          const res2 = await this._spHttpClient.post(
            `${siteUrl}/_api/web/folders/addUsingPath(decodedurl='${safeCurr}',overwrite=true)`,
            SPHttpClient.configurations.v1,
            { headers: { 'Accept': 'application/json;odata=nometadata' } }
          );
          if (res2.ok) {
            created = true;
          } else {
            lastError = await res2.text();
          }
        } catch (err: unknown) {
          lastError = err instanceof Error ? err.message : String(err);
        }
      }

      // Verify that the folder actually exists now
      const verified = await this.folderExists(siteUrl, currentPath);
      if (!verified && !created) {
        console.error(`Không thể tạo thư mục "${currentPath}":`, lastError);
        throw new Error(`Không thể tạo thư mục "${currentPath}" trên SharePoint: ${lastError}`);
      }
    }
  }

  /**
   * List files in a folder.
   */
  public async getFolderFiles(siteUrl: string, folderServerRelativeUrl: string): Promise<ISharePointFileInfo[]> {
    const safeFolder = this._escapeUrl(folderServerRelativeUrl);
    const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${safeFolder}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,Length,ETag`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Cache-Control': 'no-cache, no-store, must-revalidate',
          'Pragma': 'no-cache'
        }
      }
    );

    if (response.status === 404) {
      return [];
    }
    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`Failed to list files in folder (${response.status}): ${errText}`);
    }

    const json = await response.json();
    const items = (json.value || []) as Array<{
      Name: string;
      ServerRelativeUrl: string;
      TimeLastModified: string;
      Length: string | number;
      ETag?: string;
    }>;

    return items.map((f) => ({
      name: f.Name,
      serverRelativeUrl: f.ServerRelativeUrl,
      timeLastModified: f.TimeLastModified,
      length: typeof f.Length === 'string' ? parseInt(f.Length, 10) : f.Length,
      eTag: f.ETag
    }));
  }

  /**
   * List subfolder names in a folder.
   */
  public async getFolderSubfolderNames(siteUrl: string, folderServerRelativeUrl: string): Promise<string[]> {
    try {
      const safeFolder = this._escapeUrl(folderServerRelativeUrl);
      const endpoint = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${safeFolder}')/Folders?$select=Name`;
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        return [];
      }

      const json = await response.json();
      const items = (json.value || []) as Array<{ Name: string }>;
      return items.map((f) => f.Name).filter(Boolean);
    } catch {
      return [];
    }
  }
}
