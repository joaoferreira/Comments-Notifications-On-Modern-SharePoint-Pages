import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGetCommentsApplicationCustomizerProperties {
    testMessage: string;
}
/** A Custom Action which can be run during execution of a Client Side Application */
export default class GetCommentsApplicationCustomizer extends BaseApplicationCustomizer<IGetCommentsApplicationCustomizerProperties> {
    private _topPlaceholder;
    private _bottomPlaceholder;
    onInit(): Promise<void>;
    private getSitePageComments(id);
    private getPages();
    private UpdateItem(id, comments);
    private _renderPlaceHolders();
    private _onDispose();
}
