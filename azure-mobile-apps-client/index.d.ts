// Type definitions for Microsoft's Azure Mobile Apps: Javascript Client SDK
// Project: https://azure.github.io/azure-mobile-apps-js-client/
// Definitions by: Kelvin Dart <https://github.com/kelvindart/>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped

declare namespace Microsoft.AzureMobileAppsJS {

    interface Push {
        /**
         * Register a push channel with the Mobile Apps backend to start receiving notifications.
         *
         * @param platform The device platform being used - wns, gcm or apns.
         * @param pushChannel The push channel identifier or URI.
         * @param templates An object containing template definitions. Template objects should contain body, headers and tags properties.
         * @param secondaryTiles An object containing template definitions to be used with secondary tiles when using WNS.
         * @param callback Optional callback accepting (error, results) parameters.
         */
        register(platform: string, pushChannel: string, templates?: any, secondaryTiles?: any, callback?: (error: any, results: any) => void): void;
        /**
         * Invokes the specified custom api and returns a response object.
         *
         * @param pushChannel The push channel identifier or URI.
         * @param callback Optional callback accepting (error, results) parameters.
         */
        unregister(pushChannel: string, callback?: (error: any, results: any) => void): void;
    }

    /**
     * Represents a local store.
     *
     * API based on https://azure.github.io/azure-mobile-apps-js-client/MobileServiceStore.html
     */
    interface MobileServiceStore {
        /**
         * Defines schema of a table in the local store. If a table with the same name already exists, the newly defined
         * columns in the table definition will be added to the table. If no table with the same name exists, a table
         * with the specified schema will be created.
         *
         * @param tableDefinition An object that defines the table schema, i.e. the table name and columns.
         */
        defineTable(tableDefinition: TableDefinition): asyncPromise;

        /**
         * Deletes one or more records from the local table
         *
         * @param tableNameOrQuery Either the name of the local table in which delete is to be performed, or a QueryJs object defining records to be deleted.
         * @param ids A single ID or an array of IDs of records to be deleted. This argument is expected only if tableNameOrQuery is table name and not a QueryJs object.
         */
        del(tableNameOrQuery: string | Object, ids: string | string []): asyncPromise;

        /**
         * Perform an object / record lookup in the local table.
         *
         * @param tableName Name of the local table in which lookup is to be performed.
         * @param id id of the object to be looked up.
         * @param suppressRecordNotFoundError If set to true, lookup will return an undefined object if the record is not found. Otherwise, lookup will fail. This flag is useful to distinguish between a lookup failure due to the record not being present in the table versus a genuine failure in performing the lookup operation.
         */
        lookup(tableName: string, id: string, suppressRecordNotFoundError?: boolean);

        /**
         * Read records from a local table.
         *
         * @param query A QueryJs object representing the query to use while reading the table.
         */
        read(query: Object);

        /**
         * Updates or inserts one or more objects / records in the local table. If a property does not have a
         * corresponding definition in tableDefinition, it will not be upserted into the table.
         *
         * @param tableName Name of the local table in which data is to be upserted.
         * @param data A single object / record OR an array of objects / records to be inserted/updated in the table.
         */
        upsert(tableName: string, data: Object | Object []);
    }

    /**
     * Client for connecting to the Azure Mobile Apps backend.
     *
     * API based on https://azure.github.io/azure-mobile-apps-js-client/MobileServiceClient.html
     */
    interface MobileServiceClient {
        new (applicationUrl: string): MobileServiceClient;

        /**
         * Push registration manager.
         */
        push: Push;

        /**
         * Log a user into an Azure Mobile Apps backend.
         *
         * @param provider Name of the authentication provider to use; one of 'facebook', 'twitter', 'google', 'aad' (equivalent to 'windowsazureactivedirectory') or 'microsoftaccount'. If no provider is specified, the 'token' parameter is considered a Microsoft Account authentication token. If a provider is specified, the 'token' parameter is considered a provider-specific authentication token.
         * @param token provider specific object with existing OAuth token to log in with.
         * @param useSingleSignOn Indicates if single sign-on should be used. This parameter only applies to Windows clients and is ignored on other platforms. Single sign-on requires that the application's Package SID be registered with the Microsoft Azure Mobile Apps backend, but it provides a better experience as HTTP cookies are supported so that users do not have to login in everytime the application is launched.
         */
        login(provider: string, token: string, useSingleSignOn: boolean): asyncPromise;

        /**
         * Log a user into an Azure Mobile Apps backend.
         *
         * @param provider Name of the authentication provider to use; one of 'facebook', 'twitter', 'google', 'aad' (equivalent to 'windowsazureactivedirectory') or 'microsoftaccount'.
         * @param options Contains additional parameter information.
         */
        loginWithOptions(provider: string, options: LoginOptions): asyncPromise

        /**
         * Log a user out of the Mobile Apps backend.
         */
        logout(): asyncPromise;

        /**
         * Create a new MobileServiceClient with a filter used to process all of its network requests and responses.
         *
         * @param serviceFilter
         */
        withFilter(serviceFilter: (request: any, next: (request: any, callback: (error: any, response: any) => void) => void, callback: (error: any, response: any) => void) => void): MobileServiceClient;

        /**
         * Invokes the specified custom api and returns a response object.
         *
         * @param apiName The custom api to invoke.
         * @param options Contains additional parameter information, valid values are:
         *     body: The body of the HTTP request.
         *     method: The HTTP method to use in the request, with the default being POST,
         *     parameters: Any additional query string parameters, 
         *     headers: HTTP request headers, specified as an object.
         * @param callback Optional callback accepting (error, results) parameters.
         */
        invokeApi(apiName: string, options?: InvokeApiOptions): asyncPromise;

        /**
         * Get the associated MobileServiceSyncContext instance.
         */
        getSyncContext(): MobileServiceSyncContext;

        /**
         * Gets a reference to the specified local table.
         *
         * @param tableName
         */
        getSyncTable(tableName: string): MobileServiceSyncTable;

        /**
         * Gets a reference to the specified backend table.
         *
         * @param tableName
         */
        getTable(tableName: string): MobileServiceTable;
    }

    /**
     * A SQLite based implementation of MobileServiceStore. **Note** that this class is available _only_ as part of the
     * Cordova SDK.
     *
     * API based on https://azure.github.io/azure-mobile-apps-js-client/MobileServiceSqliteStore.html
     */
    interface MobileServiceSqliteStore extends MobileServiceStore {
        /**
         * Initializes a new instance of MobileServiceSqliteStore.
         *
         * @param dbName Name of the SQLite store file. If no name is specified, the default name will be used.
         */
        new (dbName: string): MobileServiceSqliteStore;
    }

    /**
     * Context for local store operations.
     *
     * API based on https://azure.github.io/azure-mobile-apps-js-client/MobileServiceSyncContext.html
     */
    interface MobileServiceSyncContext {
        /**
         * @param client The instance to be used to make requests to the backend (server).
         */
        new(client: MobileServiceClient): MobileServiceSyncContext;

        /**
         * Settings for configuring the pull behavior
         */
        pushHandler: IPushHandler;

        /**
         * Deletes an object / record from the specified local table.
         *
         * @param tableName Name of the local table to delete the object from.
         * @param instance The object to delete from the local table.
         */
        del(tableName: string, instance: Object): asyncPromise;

        /**
         * Initializes the MobileServiceSyncContext instance. Initailizing an initialized instance of MobileServiceSyncContext will have no effect.
         *
         * @param localStore An intitialized instance of the local store to be associated with the MobileServiceSyncContext instance.
         */
        initialize(localStore: MobileServiceStore): asyncPromise;

        /**
         * Inserts a new object / record into the specified local table. If the inserted object does not specify an id, a GUID will be used as the id.
         *
         * @param tableName Name of the local table in which the object / record is to be inserted.
         * @param instance The object / record to be inserted into the local table.
         */
        insert(tableName: string, instance: Object): asyncPromise;

        /**
         * Looks up an object / record from the specified local table using the object id.
         *
         * @param tableName Name of the local table in which to look up the object / record.
         * @param id id of the object to be looked up in the local table.
         * @param suppressRecordNotFoundError If set to true, lookup will return an undefined object if the record is not found. Otherwise, lookup will fail. This flag is useful to distinguish between a lookup failure due to the record not being present in the table versus a genuine failure in performing the lookup operation.
         */
        lookup(tableName: string, id: string, suppressRecordNotFoundError?: boolean): asyncPromise;

        /**
         * Pulls changes from server table into the local store.
         *
         * @param query Query specifying which records to pull.
         * @param queryId A unique string id for an incremental pull query. A null / undefined queryId will perform a vanilla pull, i.e. will pull all the records specified by the table from the server
         * @param settings An object that defines various pull settings.
         */
        pull(query, queryId, settings): asyncPromise;

        /**
         * Purges data from the local table as well as pending operations and any incremental sync state associated with
         * the table.
         *
         * A regular purge, would fail if there are any pending operations for the table being purged.
         *
         * A forced purge will proceed even if pending operations for the table being purged exist in the operation
         * table. In addition, it will also delete the table's pending operations.
         *
         * @param query A QueryJs object representing the query that specifies what records are to be purged.
         * @param forcePurge If set to true, the method will perform a forced purge.
         */
        purge(query, forcePurge): asyncPromise;

        /**
         * Pushes local changes to the corresponding tables on the server.
         *
         * Conflict and error handling are delegated to MobileServiceSyncContext#pushHandler.
         */
        push(): asyncPromise;

        /**
         * Reads records from the specified local table.
         *
         * @param query A QueryJs object representing the query to use while reading the local table
         */
        read(query: Object): asyncPromise;

        /**
         * Update an object / record in the specified local table. The id of the object / record identifies the record that will be updated in the table.
         *
         * @param tableName Name of the local table in which the object / record is to be updated.
         * @param instance New value of the object / record to be updated.
         */
        update(tableName: string, instance: Object): asyncPromise;
    }

    /**
     * Represents a table in the local store.
     *
     * API based on https://azure.github.io/azure-mobile-apps-js-client/MobileServiceSyncTable.html
     */
    interface MobileServiceSyncTable {
        /**
         * @param tableName Name of the local table.
         * @param client The MobileServiceClient instance associated with this table.
         */
        new (tableName: string, client: MobileServiceClient): MobileServiceSyncTable;

        /**
         * Deletes an object / record from the local table.
         *
         * @param instance The object to delete from the local table.
         */
        del(instance: Object): asyncPromise;

        /**
         * Gets the MobileServiceClient instance associated with this table.
         */
        getMobileServiceClient(): MobileServiceClient;

        /**
         * Gets the name of the backend table.
         */
        getTableName(): string;

        /**
         * Inserts a new object / record in the local table. If the inserted object does not specify an id, a GUID will
         * be used as the id.
         *
         * @param instance
         */
        insert(instance: Object): asyncPromise;

        /**
         * Looks up an object / record from the local table using the object id.
         *
         * @param id id of the object to be looked up in the local table.
         * @param suppressRecordNotFoundError If set to true, lookup will return an undefined object if the record is not found. Otherwise, lookup will fail. This flag is useful to distinguish between a lookup failure due to the record not being present in the table versus a genuine failure in performing the lookup operation.
         */
        lookup(id: string, suppressRecordNotFoundError: boolean): asyncPromise;

        /**
         * Pulls changes from server table into the local table.
         *
         * @param query Query specifying which records to pull.
         * @param queryId A unique string id for an incremental pull query. A null / undefined queryId will perform a vanilla pull, i.e. will pull all the records specified by the table from the server
         * @param settings An object that defines various pull settings.
         */
        pull(query: Object, queryId: string, settings: IPullSettings): asyncPromise;

        /**
         * Purges data from the local table as well as pending operations and any incremental sync state associated with
         * the table.
         *
         * A regular purge, would fail if there are any pending operations for the table being purged.
         *
         * A forced purge will proceed even if pending operations for the table being purged exist in the operation
         * table. In addition, it will also delete the table's pending operations.
         *
         * @param query A QueryJs object representing the query that specifies what records are to be purged.
         * @param forcePurge If set to true, the method will perform a forced purge.
         */
        purge(query: Object, forcePurge: boolean): asyncPromise;

        /**
         * Reads records from the local table.
         *
         * @param query A QueryJs object representing the query to use while reading the local table
         */
        read(query: Object): asyncPromise;

        /**
         * Update an object / record in the local table.
         *
         * @param instance New value of the object / record.
         */
        update(instance: Object): asyncPromise;
    }

    /**
     * Defines callbacks for performing conflict and error handling.
     *
     * API based on https://azure.github.io/azure-mobile-apps-js-client/MobileServiceSyncContext.html#.PushHandler
     */
    interface IPushHandler {
        /**
         *  Callback for delegating conflict handling to the user.
         *
         *  @param pushError The pushError object raised for current conflict
         */
        onConflict(pushError: IPushError): Promise | undefined;

        /**
         *  Callback for delegating error handling to the user.
         *
         *  @param pushError The pushError object raised for current error
         */

        onError(pushError: IPushError): Promise | undefined;
    }

    /**
     * A conflict / error encountered while pushing a change to the server. This wraps the underlying error and provides
     * additional helper methods for handling the conflict / error.
     *
     * API based on https://azure.github.io/azure-mobile-apps-js-client/PushError.html
     */
    interface IPushError {
        new(): IPushError;

        /**
         * When set to true, the push logic will consider the conflict / error to be handled. In such a case, an attempt would be made to push the change to the server again. By default, isHandled is set to false.
         */
        isHandled: boolean;

        /**
         * Cancels pushing the current operation to the server permanently. This will also set PushError#isHandled to true.
         *
         * This method simply removes the pending operation from the operation table, thereby permanently skipping the
         * associated change from being pushed to the server. A future change done to the same record will not be
         * affected and such changes will continue to be pushed.
         *
         * @return A promise that is fulfilled when the operation is cancelled OR rejected with the error if it fails to cancel it.
         */
        cancel(): asyncPromise;

        /**
         * Cancels the push operation for the current record and discards the record from the local store.
         *
         * This will also set PushError#isHandled to true.
         *
         * @return A promise that is fulfilled when the operation is cancelled and the client record is discarded and rejected with the error if cancelAndDiscard fails.
         */
        cancelAndDiscard(): asyncPromise;

        /**
         * Cancels the push operation for the current record and updates the record in the local store.
         *
         * This will also set PushError#isHandled to true.
         *
         * @param newValue
         *
         * @return A promise that is fulfilled when the operation is cancelled and the client record is updated. The promise is rejected with the error if cancelAndUpdate fails.
         */
        cancelAndUpdate(newValue: Object): asyncPromise;

        /**
         * Changes the type of operation that will be pushed to the server.
         * This is useful for handling conflicts where you might need to change the type of the
         * operation to be able to push the changes to the server.
         * This will also set {@link PushError#isHandled} to true.
         *
         * Example: You might need to change _'insert'_ to _'update'_ to be able to push a record that
         * was already inserted on the server.
         *
         * Note: Changing the action to _'delete'_' will automatically remove the associated record from the
         * data table in the local store.
         *
         * @param newAction New type of the operation. Valid values are _'insert'_, _'update'_ and _'delete'_.
         * @param newClientRecord New value of the client record.
         *                        The `id` property of the new record should match the `id` property of the original record.
         *                        If `newAction` is _'delete'_, only the `id` and `version` properties will be read from `newClientRecord`.
         *                        Reading the `version` property while deleting is useful if
         *                        the conflict handler changes an _'insert'_  /_'update'_ action to _'delete'_' and also updated the version
         *                        to reflect the server version.
         */
        changeAction(newAction: string, newClientRecord: Object): asyncPromise;

        /**
         * Gets the action for which conflict / error occurred.
         */
        getAction(): string;

        /**
         * Gets the value of the record that was pushed to the server when the conflict /error occurred. Note that this
         * may not be the latest value as local tables could have changed after we started the push operation.
         */
        getClientRecord(): Object;

        /**
         * Gets the underlying error encountered while performing the push operation. This contains grannular details of
         * the failure like server response, error code, etc.
         *
         * Note: Modifying value returned by this method will have a side effect of permanently changing the underlying
         * error object
         */
        getError(): Error;

        /**
         * Gets the value of the record on the server, if available, when the conflict / error occurred. This is useful
         * while handling conflicts. However, note that in the event of a conflict / error, the server may not always
         * respond with the server record's value. Example: If the push failed even before the client value reaches the
         * server, we won't have the server value. Also, there are some scenarios where the server does not respond with
         * the server value.
         */
        getServerRecord(): Object;

        /**
         * Gets the name of the table for which conflict / error occurred.
         */
        getTableName(): string;

        /**
         * Checks if the error is a conflict.
         */
        isConflict(): boolean;

        /**
         * Updates the client data record associated with the current operation. If required, the metadata in the log
         * record will also be associated. This will also set PushError#isHandled to true.
         *
         * @param newValue New value of the data record.
         */
        update(newValue): asyncPromise;
    }

    /**
     * API based on https://azure.github.io/azure-mobile-apps-js-client/MobileServiceSyncContext.html#.PullSettings
     */
    interface IPullSettings {
        pageSize?: number
    }

    /**
     * An object that defines the table schema, i.e. the table name and columns.
     */
    interface TableDefinition {
        name?: string,
        columnDefinitions?: Object
    }

    interface InvokeApiOptions {
		method?: string;
		body?: any;
		headers?: Object;
		parameters?: Object;
	}

    interface LoginOptions {
        token?: Object;
        useSingleSignOn?: boolean;
        parameters?: Object
    }

    /**
     * MobileServiceTable object based on Microsoft Azure documentation: http://msdn.microsoft.com/en-us/library/windowsazure/jj554239.aspx
     */
    interface MobileServiceTable extends IQuery {
        new (tableName: string, client: MobileServiceClient): MobileServiceTable;

        getTableName(): string;

        getMobileServiceClient(): MobileServiceClient;

        insert(instance: any, paramsQS: Object): asyncPromise;

        update(instance: any, paramsQS: Object): asyncPromise;

        lookup(id: number, paramsQS: Object): asyncPromise;

        del(instance: any, paramsQS: Object): asyncPromise;

        read(query: IQuery, paramsQS: Object): asyncPromise;
    }


    /**
     * Interface to describe Query object fluent creation based on Microsoft Azure documentation: http://msdn.microsoft.com/en-us/library/windowsazure/jj613353.aspx
     */
    interface IQuery {
        read(paramsQS?: Object): asyncPromise;

        orderBy(...propName: string[]): IQuery;
        orderByDescending(...propName: string[]): IQuery;
        select(...propNameSelected: string[]): IQuery;
        select(funcProjectionFromThis: () => any): IQuery;
        where(mapObjFilterCriteria: any): IQuery;
        where(funcPredicateOnThis: (...qParams: any[]) => boolean, ...qValues: any[]): IQuery;
        skip(n: number): IQuery;
        take(n: number): IQuery;
        includeTotalCount(): IQuery;
    }

    /**
     * Interface to Platform.async(func) => Platform.Promise based on code MobileServices.Web-1.0.0.js
     */
    interface asyncPromise {
        then(onSuccess: (result: any) => any, onError?: (error: any) => any): asyncPromise;
        done(onSuccess?: (result: any) => void , onError?: (error: any) => void ): void;
    }

    interface AzureMobileAppsJS {
        MobileServiceClient: MobileServiceClient;
    }
}

declare var AzureMobileAppsJS: Microsoft.AzureMobileAppsJS.AzureMobileAppsJS;
