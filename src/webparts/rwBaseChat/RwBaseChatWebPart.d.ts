import "@pnp/sp/lists";
import "@pnp/sp/webs";
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface IRwBaseChatWebPartProps {
    list: string;
    title: string;
    background: string;
    interval: any;
    autofetch: boolean;
}
export interface ISPLists {
    value: ISPList[];
}
export interface ISPList {
    Id: string;
    Zeit: string;
    User: string;
    Message: string;
}
export default class RwBaseChatWebPart extends BaseClientSideWebPart<IRwBaseChatWebPartProps> {
    protected onInit(): Promise<void>;
    checkChatList(list: string): Promise<void>;
    private _getLatestMessageTime;
    private _getNewMessages;
    render(): void;
    innerEmoji(emoji: string): void;
    private debounce;
    _getMessage(): Promise<void>;
    _addToList(zeit: string, user: string, message: string): Promise<void>;
    checkListCreation(listName: string): void;
    _setFieldTypes(): Promise<void>;
    private _renderListAsync;
    private _renderList;
    _getListData(): Promise<ISPLists>;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=RwBaseChatWebPart.d.ts.map