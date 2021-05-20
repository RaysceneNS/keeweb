/* eslint-disable no-console */
import { StorageBase } from 'storage/storage-base';

class StorageAzure extends StorageBase {
    name = 'azure';
    enabled = true;
    uipos = 50;
    icon = 'cube';

    getPathForName(fileName) {
        return '/' + fileName + '.kdbx';
    }

    load(path, opts, callback) {
        this._oauthAuthorize((err) => {
            if (err) {
                return callback && callback(err);
            }
            this.logger.debug('Load', path);
            const ts = this.logger.ts();
            const url = this._blobStoreUrl(path);
            this.logger.debug('Load url', url);
            this._xhr({
                url,
                headers: {
                    'x-ms-date': new Date().toGMTString(),
                    'x-ms-version': '2020-06-12'
                },
                responseType: 'arraybuffer',
                success: (response, xhr) => {
                    const rev = xhr.getResponseHeader('ETag');
                    this.logger.debug('Loaded', path, rev, this.logger.ts(ts));
                    return callback && callback(null, response, { rev });
                },
                error: (err) => {
                    this.logger.error('Load error', path, err, this.logger.ts(ts));
                    return callback && callback(err);
                }
            });
        });
    }

    stat(path, opts, callback) {
        this._oauthAuthorize((err) => {
            if (err) {
                return callback && callback(err);
            }
            this.logger.debug('Stat', path);
            const ts = this.logger.ts();
            const url = this._blobStoreUrl(path);
            this._xhr({
                url,
                headers: {
                    'x-ms-date': new Date().toGMTString(),
                    'x-ms-version': '2020-06-12'
                },
                method: 'HEAD',
                success: (response, xhr) => {
                    const rev = xhr.getResponseHeader('ETag');
                    if (!rev) {
                        this.logger.error('Stat error', path, 'no eTag', this.logger.ts(ts));
                        return callback && callback('no eTag');
                    }
                    this.logger.debug('Stated', path, rev, this.logger.ts(ts));
                    return callback && callback(null, { rev });
                },
                error: (err, xhr) => {
                    if (xhr.status === 404) {
                        this.logger.debug('Stated not found', path, this.logger.ts(ts));
                        return callback && callback({ notFound: true });
                    }
                    this.logger.error('Stat error', path, err, this.logger.ts(ts));
                    return callback && callback(err);
                }
            });
        });
    }

    save(path, opts, data, callback, rev) {
        this._oauthAuthorize((err) => {
            if (err) {
                return callback && callback(err);
            }
            this.logger.debug('Save', path, rev);
            const ts = this.logger.ts();
            const url = this._blobStoreUrl(path);
            this._xhr({
                url,
                headers: {
                    'x-ms-date': new Date().toGMTString(),
                    'x-ms-version': '2020-06-12',
                    'x-ms-blob-type': 'BlockBlob',
                    'If-Match': rev
                },
                method: 'PUT',
                data,
                statuses: [201, 409, 412],
                success: (response, xhr) => {
                    rev = xhr.getResponseHeader('ETag');
                    if (!rev) {
                        this.logger.error('Save error', path, 'no eTag', this.logger.ts(ts));
                        return callback && callback('no eTag');
                    }
                    if (xhr.status === 409) {
                        this.logger.debug('Save error', path, rev, this.logger.ts(ts));
                        return callback && callback({ revConflict: true }, { rev });
                    }
                    if (xhr.status === 412) {
                        this.logger.debug('Save conflict', path, rev, this.logger.ts(ts));
                        return callback && callback({ revConflict: true }, { rev });
                    }
                    this.logger.debug('Saved', path, rev, this.logger.ts(ts));
                    return callback && callback(null, { rev });
                },
                error: (err) => {
                    this.logger.error('Save error', path, err, this.logger.ts(ts));
                    return callback && callback(err);
                }
            });
        });
    }

    list(dir, callback) {
        this._oauthAuthorize((err) => {
            if (err) {
                return callback && callback(err);
            }
            this.logger.debug('List', dir);
            const ts = this.logger.ts();
            const isRoot = !dir || dir.length === 0;
            const url = isRoot
                ? this._blobStoreUrl('?restype=container&comp=list')
                : this._blobStoreUrl(`${dir}?restype=container&comp=list`);

            this._xhr({
                url,
                headers: {
                    'x-ms-date': new Date().toGMTString(),
                    'x-ms-version': '2020-06-12'
                },
                responseType: 'document',
                success: (response) => {
                    if (!response) {
                        this.logger.error('List error', this.logger.ts(ts), response);
                        return callback && callback('list error');
                    }
                    const fileList = [];
                    // treat containers as directories
                    response.documentElement.querySelectorAll('Container').forEach((blob) => {
                        const name = blob.getElementsByTagName('Name')[0].textContent;
                        const etag = blob.getElementsByTagName('Etag')[0].textContent;
                        fileList.push({
                            path: name,
                            name,
                            dir: true,
                            rev: etag
                        });
                    });
                    // treat blobs as files
                    response.documentElement.querySelectorAll('Blob').forEach((blob) => {
                        const name = blob.getElementsByTagName('Name')[0].textContent;
                        const etag = blob.getElementsByTagName('Etag')[0].textContent;
                        fileList.push({
                            path: isRoot ? name : dir + '/' + name,
                            name,
                            dir: false,
                            rev: etag
                        });
                    });
                    this.logger.debug('Listed', this.logger.ts(ts), fileList);
                    return callback && callback(null, fileList);
                },
                error: (err) => {
                    this.logger.error('List error', this.logger.ts(ts), err);
                    return callback && callback(err);
                }
            });
        });
    }

    remove(path, callback) {
        this.logger.debug('Remove', path);
        const ts = this.logger.ts();
        const url = this._blobStoreUrl(path);
        this._xhr({
            url,
            headers: {
                'x-ms-date': new Date().toGMTString(),
                'x-ms-version': '2020-06-12'
            },
            method: 'DELETE',
            statuses: [202, 204],
            success: () => {
                this.logger.debug('Removed', path, this.logger.ts(ts));
                return callback && callback();
            },
            error: (err) => {
                this.logger.error('Remove error', path, err, this.logger.ts(ts));
                return callback && callback(err);
            }
        });
    }

    logout(enabled) {
        this._oauthRevokeToken();
    }

    _blobStoreUrl(path) {
        const basePath = new URL(this.appSettings.azureBlobContainer).toString();
        return new URL(path, basePath).toString();
    }

    _getOAuthConfig() {
        const clientId = this.appSettings.azureClientId;
        const tenantid = this.appSettings.azureTenantId;
        return {
            url: `https://login.microsoftonline.com/${tenantid}/oauth2/v2.0/authorize`,
            tokenUrl: `https://login.microsoftonline.com/${tenantid}/oauth2/v2.0/token`,
            scope: 'https://storage.azure.com/user_impersonation',
            clientId,
            pkce: true,
            width: 600,
            height: 500
        };
    }
}

export { StorageAzure };
