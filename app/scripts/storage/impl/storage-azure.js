/* eslint-disable no-console */
import { StorageBase } from 'storage/storage-base';
import { AzureApps } from 'const/cloud-storage-apps';
import { Features } from 'util/features';

class StorageAzure extends StorageBase {
    name = 'azure';
    enabled = true;
    uipos = 50;
    icon = 'cube';

    _baseUrl = 'https://fskeepasstest.blob.core.windows.net/';

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
            const url = this._baseUrl + path;
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
            const url = this._baseUrl + path;
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
            const url = this._baseUrl + path;
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
                ? this._baseUrl + '?comp=list'
                : this._baseUrl + `${dir}?restype=container&comp=list`;

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
                    response.documentElement
                        .querySelectorAll(isRoot ? 'Container' : 'Blob')
                        .forEach((blob) => {
                            const name = blob.getElementsByTagName('Name')[0].textContent;
                            const etag = blob.getElementsByTagName('Etag')[0].textContent;
                            fileList.push({
                                path: isRoot ? name : dir + '/' + name,
                                name,
                                dir: isRoot,
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
        const url = this._baseUrl + path;
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

    _getOAuthConfig() {
        let clientId = this.appSettings.azureClientId;
        let clientSecret = this.appSettings.azureClientSecret;
        if (!clientId) {
            if (Features.isDesktop) {
                ({ id: clientId, secret: clientSecret } = AzureApps.Desktop);
            } else if (Features.isLocal) {
                ({ id: clientId, secret: clientSecret } = AzureApps.Local);
            } else {
                ({ id: clientId, secret: clientSecret } = AzureApps.Production);
            }
        }

        return {
            url:
                'https://login.microsoftonline.com/3c48f4df-462c-4439-899d-9ed41948d939/oauth2/v2.0/authorize',
            tokenUrl:
                'https://login.microsoftonline.com/3c48f4df-462c-4439-899d-9ed41948d939/oauth2/v2.0/token',
            scope: 'https://storage.azure.com/user_impersonation',
            clientId,
            clientSecret,
            pkce: true,
            width: 600,
            height: 500
        };
    }
}

export { StorageAzure };
