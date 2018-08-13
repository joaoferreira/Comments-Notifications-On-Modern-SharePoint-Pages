import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName 
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'GetCommentsApplicationCustomizerStrings';

import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';
import pnp from "sp-pnp-js";
import { PermissionKind } from "@pnp/sp";

const LOG_SOURCE: string = 'GetCommentsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGetCommentsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GetCommentsApplicationCustomizer
  extends BaseApplicationCustomizer<IGetCommentsApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {


    pnp.setup({
      spfxContext: this.context
    });
      
    pnp.sp.web.currentUserHasPermissions(PermissionKind.ManageWeb).then(perms => {
      if(perms){
        this._renderPlaceHolders();
      }      
    });
    //this.getSitePageComments();
    //this.getPages();
    

    return Promise.resolve();
  }

  private async getSitePageComments(id) {
      const currentWebUrl: string = this.context.pageContext.web.serverRelativeUrl;

      const response = await this.context.spHttpClient.get(`${currentWebUrl}/_api/web/lists/GetByTitle('Site Pages')/GetItemById(${id})/Comments?$expand=replies,likedBy,replies/likedBy&$top=10&$inlineCount=AllPages`, SPHttpClient.configurations.v1);

      const responseJSON = await response.json();

      let commentsNumber: number = 0;

      for (let entry of responseJSON.value) {
        commentsNumber ++;
        commentsNumber = commentsNumber + entry.replyCount;      
      } 
      this.UpdateItem(id, commentsNumber);
      console.log('comments: ' +commentsNumber + 'page id: ' + id); 
  }


  private async getPages(){
    const currentWebUrl: string = this.context.pageContext.web.serverRelativeUrl;

    const response = await this.context.spHttpClient.get(`${currentWebUrl}/_api/web/lists/GetByTitle('Site Pages')/items`, SPHttpClient.configurations.v1);

    const responseJSON = await response.json();
    let pageId;
    for (let entry of responseJSON.value) {
      pageId = entry.ID;
      
     
      this.getSitePageComments(pageId);
    }
    alert('SharePoint is processing the page comments!')
  }


  private UpdateItem(id, comments)
  {
    pnp.sp.web.lists.getByTitle("Site Pages").items.getById(id).update({
      Previous_x0020_Comments : comments
    }).then(console.log)
    .catch(console.log);
  }



  private _renderPlaceHolders(): void {


      // Handling the bottom placeholder
      if (!this._bottomPlaceholder) {
        this._bottomPlaceholder =
          this.context.placeholderProvider.tryCreateContent(
            PlaceholderName.Bottom,
            { onDispose: this._onDispose });
  
        // The extension should not assume that the expected placeholder is available.
        if (!this._bottomPlaceholder) {
          console.error('The expected placeholder (Bottom) was not found.');
          return;
        }
  
        if (this.properties) {
  
          if (this._bottomPlaceholder.domElement) {
            this._bottomPlaceholder.domElement.innerHTML = `
              <div id="checkComments" Title="Get Comments Notfications" style="position: absolute; bottom: 0; width: 22px; height: 18px; left: 10px; z-index: 100; padding: 10px; cursor: pointer;" class="ms-bgColor-themeDark ms-fontColor-white ">
                <i class="ms-Icon ms-Icon--Message" aria-hidden="true" style="font-size: 20px;"></i>
              </div>`;
          }
        }
      }
      let ctx = this;
      document.getElementById('checkComments').onclick = function(){ctx.getPages()};

    }

    private _onDispose(): void {
      console.log('Disposed Coments.');
    }





}
