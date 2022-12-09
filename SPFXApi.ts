import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import {
    SPHttpClient,
    SPHttpClientConfiguration,
    SPHttpClientResponse
} from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base'; 

export interface IListProperties {
    listId: string;
    select?: string[];
    filter?: string[];
    expand?: string[];
    orderBy?: string[];
    top?: string;
}

export interface IAPIResult {
    message: string;
    value: any[];
}

export default class APIServices {
    // private property
    private siteURL: string;    // Site URL 
    private context: any; 
    
    // Factory api constructor with siteURL parameter
    public constructor( ctx?:ApplicationCustomizerContext | WebPartContext | undefined){
        if(ctx) this.context = ctx;
        this.siteURL = this.context.pageContext.web.absoluteUrl;
    }

    private bindSelector = (url: string, key: string, value: string[]) => {
        return url + (key ? ((url.includes("?$") ? `&$${key}=` : `?$${key}=`) + value.toString()) : '');
    }

    // Get all available list details
    public getAllLists = async (props: { select?:string []}): Promise<{ value: any }> => { 
        try {
            const {select} = props ? props : {select : []};  
            const response = await this.context.spHttpClient.get(
                this.bindSelector(this.siteURL + `/_api/web/lists?$filter=Hidden eq false`, "Select", select), 
                SPHttpClient.configurations.v1,
                {
                    headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                    }
                }
            );
            return response.json();
        } catch (e) {
            console.log("Something went wrong!", e);
            throw e;
        }
    }

    // Get all list items using list properties 
    public getAllFields = async (props: {listId: string, select: string[], filter: string[]}): Promise<any> => {  
        try {
            const {listId, select, filter} = props;  
            let url = `${this.siteURL}/_api/lists(guid'${listId}')/fields`; 
            if(select && select.length>0) 
                url = this.bindSelector(url, "Select", select);
            if(filter && filter.length>0) 
                url = this.bindSelector(url, "filter", filter); 
            const data = await this.context.spHttpClient.get(url, 
                SPHttpClient.configurations.v1,
                {
                    headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                    }
                }
            );
            return data.json();
        } catch (e) {
            console.log("Something went wrong!", e);
            throw e;
        } 
    }

    // Get all list items using list properties 
    public getAllItems = async (props: {listId: string, select?: string[], filter?: string[], expand?: string[], orderBy?: string[]}): Promise<IAPIResult> => {  
        try {
            const {listId, select, filter, expand, orderBy} = props;  
            let url = `${this.siteURL}/_api/lists(guid'${listId}')/items`; 
            if(select && select.length>0) url = this.bindSelector(url, "Select", select);
            if(filter && filter.length>0) url = this.bindSelector(url, "filter", filter); 
            if(expand && expand.length>0) url = this.bindSelector(url, "expand", expand); 
            if(orderBy && orderBy.length>0) url = this.bindSelector(url, "orderBy", orderBy); 
            const data = await this.context.spHttpClient.get(url, 
                SPHttpClient.configurations.v1,
                {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'odata-version': ''
                    }
                }
            );
            return data.json();
        }  catch (e) {
            console.log("Something went wrong!", e);
            throw e;
        } 
    }

    // Get all list items using list properties and item ID
    public getItemById = async (props: {listId: number, itemID: number, select: string[], filter: string[], expand: string[], orderBy: string[]}): Promise<IAPIResult> => {  
        try {
            const {listId, select, filter, expand, orderBy, itemID} = props;  
            let url = `${this.siteURL}/_api/lists(guid'${listId}')/items(${itemID})`; 
            if(select && select.length>0) url = this.bindSelector(url, "Select", select);
            if(filter && filter.length>0) url = this.bindSelector(url, "filter", filter); 
            if(expand && expand.length>0) url = this.bindSelector(url, "expand", expand); 
            if(orderBy && orderBy.length>0) url = this.bindSelector(url, "orderBy", orderBy); 
            const data = await this.context.spHttpClient.get(url, 
                SPHttpClient.configurations.v1,
                {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'odata-version': ''
                    }
                }
            );
            return data.json();
        }  catch (e) {
            console.log("Something went wrong!", e);
            throw e;
        } 
    }

    // Get all list items using list properties 
    public getListAttachment = async (props: {listName: string, itemID: number}): Promise<IAPIResult> => {  
        const url = this.siteURL + `/_api/web/lists/getByTitle('${props.listName}')/items(${props.itemID})/AttachmentFiles`;
        return new Promise<any>((resolve: any, reject: (error: any)=> void): void => {  
            this.context.spHttpClient.get(url, 
                SPHttpClient.configurations.v1,
                {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'odata-version': ''
                    }
                }
            ).then((response:SPHttpClientResponse)=>{
                resolve(response.json());
            }).catch(err=>{
                reject(err);
            });
        });
    }

    // Create list using item using list properties 
    public createItem = (listName: string, body: object): Promise<IAPIResult> => { 
        return new Promise<any>((resolve: any, reject: (error: any)=> void): void => {  
            const Url = `${this.siteURL}/_api/web/lists/getbytitle('${listName}')/items`;
            this.context.spHttpClient.post(Url,
            SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=nometadata',
                    'odata-version': ''
                },
                body: JSON.stringify(body)
            }).then((response: SPHttpClientResponse)=>{
                resolve(response.json());
            }).catch((err)=>{
                console.log(err.message);
                reject(err);
            });
        });   
    }

    // Upload list attachment
    public updateItem = async (
        ID: number, 
        listName: string, 
        body: any
    ): Promise<IAPIResult>=> {
        const Url = `${this.siteURL}/_api/web/lists/getbytitle('${listName}')/items(${ID})`; 
        return await this.context.spHttpClient.post(Url,
            SPHttpClient.configurations.v1, {
                headers: { 
                    "Accept": "application/json;odata=verbose",
                    "Content-type": "application/json;odata=verbose",
                    "odata-version": "",
                    "IF-MATCH": "*",
                    "X-HTTP-Method": "MERGE"
                },
                body: JSON.stringify(body),
            }).then(async(response: SPHttpClientResponse)=>{
                return await response.json();
            }).catch(async(err)=>{
                console.log(err.message);
                return await err;
            });
    }

    // Delete list item
    public deleteItem = async (ID: number, listName: string): Promise<IAPIResult>=> { 
        const Url = `${this.siteURL}/_api/web/lists/getbytitle('${listName}')/items(${ID})`;
        return await this.context.spHttpClient.post(Url,
            SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=nometadata',
                    'odata-version': '',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE'
                }
            }).then(async(response: SPHttpClientResponse)=>{
                return await response.json();
            }).catch(async(err)=>{
                console.log(err.message);
                return await err;
            });
    }

    // Delete list item attachment
    public deleteItemAttachment = async (ID: number, listName: string, fileName: string): Promise<IAPIResult>=> { 
        const Url = `${this.siteURL}/_api/web/lists/getbytitle('${listName}')/GetItemById(${ID})/AttachmentFiles/getByFileName('${fileName}')`;
        return await this.context.spHttpClient.post(Url,
            SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=nometadata',
                    'odata-version': '',
                    'IF-MATCH': '*',
                    'X-HTTP-Method': 'DELETE'
                }
            }).then(async(response: SPHttpClientResponse)=>{
                return await response.json();
            }).catch(async(err)=>{
                console.log(err.message);
                return await err;
            });
    }

    // Upload list attachment
    public uploadListAttachment = async (
        ID: number, 
        listName: string, 
        file: any, 
        buffer: any
    ): Promise<any>=> {
        const Url = `${this.siteURL}/_api/web/lists/getbytitle('${listName}')/items(${ID})/AttachmentFiles/add(FileName='${file.name}')`;
        let resp;
        await this.context.spHttpClient.post(Url,
            SPHttpClient.configurations.v1, {
                headers: { 
                    "Content-type": "application/json;odata=verbose"
                },
                body: buffer,
            }).then( async (response: SPHttpClientResponse) => {
                await response.json().then((responseJSON: JSON) => {
                    resp = responseJSON;
                });
            }).catch(err=>{
                resp = err;
            }); 
        return resp; 
    }

    public getFileBuffer = (file: any) => {
        var reader = new FileReader();
        return new Promise<any>((resolve: any, reject: (error: any)=> void): void => {  
            reader.onload = (e) => {
                resolve(e.target.result);
            };
            reader.onerror = (e) => {
                reject(e.target.error);
            };
            reader.readAsArrayBuffer(file);
        });
    }  
}
