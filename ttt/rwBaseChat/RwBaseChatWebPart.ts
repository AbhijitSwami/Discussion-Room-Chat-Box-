import {  spfi, SPFx} from "@pnp/sp";  // Correct import
import "@pnp/sp/lists";
import "@pnp/sp/webs"; 



import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './RwBaseChatWebPart.module.scss';
import * as strings from 'RwBaseChatWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPost } from "./SPOHelper";

// Declare sp at the class level
//let sp: SPFI;

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

let listName: string;

export default class RwBaseChatWebPart extends BaseClientSideWebPart<IRwBaseChatWebPartProps> {

  protected async onInit(): Promise<void> {
    await super.onInit();

    // Initialize PnP SP object inside onInit

    if (this.properties.autofetch) {
      setInterval(() => this.checkChatList(this.properties.list), this.properties.interval);
    }
  }


  public async checkChatList(list: string): Promise<void> {
    try {
      const latestMessageTime = this._getLatestMessageTime();
      const response = await this._getNewMessages(latestMessageTime);

      if (response.value.length) {
        this._renderList(response.value);
      }
    } catch (error) {
      console.error("Error fetching messages:", error);
    }
  }

  private _getLatestMessageTime(): string {
    const messagesContainer = this.domElement.querySelectorAll('.bubble');
    const lastMessageElement = messagesContainer[messagesContainer.length - 1] as HTMLElement;
    return lastMessageElement.dataset.timestamp || '';
  }

  private async _getNewMessages(timestamp: string): Promise<ISPLists> {
    const filter = timestamp ? `?$filter=Zeit gt '${timestamp}'` : '';
    const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items${filter}`;
    
    const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    return response.json();
  }

  public render(): void {
    listName = this.properties.list;
    this.domElement.innerHTML = `
      <div class="${styles.rwBaseChat}">
        <div class="${styles.container}">
          <div class="${styles.row}" style="background-color: ${this.properties.background};">
            <span class="${styles.title}">${escape(this.properties.title)}</span><br><br>
            <div id="spListContainer" style="max-height: 300px; overflow: auto; background-color: ${this.properties.background}"></div>
          </div>
          <div class="${styles.row}" style="background-color: ${this.properties.background};">
            <table style="width: 100%; text-align: center;">
              <tr>
                <td><button id="emoji" class="${styles.button2}"><span>&#128521;</span></button></td>
                <td><textarea id="message" placeholder="Your Message" class="${styles.textarea}"></textarea></td>
                <td><button id="submit" class="${styles.button2}"><span>&#9993;</span></button></td>
              </tr>
            </table>
            <div id="emojibar" class="${styles.emojibar}">
              <button class="${styles.emoji}">😂</button>
              <button class="${styles.emoji}">😅</button>
              <button class="${styles.emoji}">😉</button>
              <button class="${styles.emoji}">😇</button>
              <button class="${styles.emoji}">😍</button>
              <!-- More emojis -->
            </div>
          </div>
        </div>
      </div>`;

    // Add event listeners
let btnSubmit = this.domElement.querySelector("#submit");
if (btnSubmit) { // Check if btnSubmit is not null
  btnSubmit.addEventListener("click", this.debounce(() => this._getMessage(), 300));
}

let btnEmojiButton = this.domElement.querySelectorAll("." + styles.emoji);
btnEmojiButton.forEach((item) => {
  const emoji = item.textContent || ''; // Provide a fallback value if textContent is null
  item.addEventListener("click", () => this.innerEmoji(emoji));
});


let btnEmoji = this.domElement.querySelector("#emoji");
if (btnEmoji) { // Check if btnEmoji is not null
  btnEmoji.addEventListener("click", () => {
    const emojiBar = this.domElement.querySelector("#emojibar");
    if (emojiBar) { // Check if emojiBar is not null
      emojiBar.classList.toggle(styles.emojibarOpen);
      emojiBar.classList.toggle(styles.emojibar);
    }
  });
}

    this.checkListCreation(this.properties.list);
    this._renderListAsync();
  }

  public innerEmoji(emoji: string): void {
    let txaMessage = this.domElement.querySelector('textarea') as HTMLTextAreaElement | null; // Cast to HTMLTextAreaElement
    if (txaMessage) { // Check if txaMessage is not null
      txaMessage.value += emoji; // Append the emoji to the textarea value
    }
  }

  private debounce(fn: (...args: any[]) => void, delay: number): (this: any, ...args: any[]) => void {
    let timeoutID: ReturnType<typeof setTimeout>; // Use ReturnType to infer the type of timeoutID
    return function (this: any, ...args: any[]) {
      if (timeoutID) clearTimeout(timeoutID);
      timeoutID = setTimeout(() => fn.apply(this, args), delay);
    };
  }

  public async _getMessage(): Promise<void> {
    try {
      let user = this.context.pageContext.user.displayName || this.context.pageContext.user.loginName;
      let messageElement = this.domElement.querySelector("textarea") as HTMLTextAreaElement | null; // Cast to HTMLTextAreaElement
      if (messageElement) { // Check if messageElement is not null
        let message = messageElement.value; // Now it's safe to access value
        let time = new Date().toLocaleString();
  
        await this._addToList(time, user, message);
        messageElement.value = ''; // Clear the input
        await this._renderListAsync();
      } else {
        console.warn("Textarea element not found.");
      }
    } catch (error) {
      console.error("Error submitting message:", error);
    }
  }

  public async _addToList(zeit: string, user: string, message: string): Promise<void> {
    try {
      await SPPost({
        url: `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
        payload: { Zeit: zeit, User: user, Message: message }
      });
    } catch (error) {
      console.error("Error adding message to list:", error);
    }
  }
  

  public checkListCreation(listName: string): void {
    // Initialize PnP SP object in context
    const sp = spfi().using(SPFx(this.context)); // Use spfi with SPFx context

    // Ensure list creation
    sp.web.lists.ensure(listName).then(async ({ created, list }) => {
        if (created) {
            this._setFieldTypes(); // Call your method to set field types
            console.log(`List ${listName} was created.`);
        } else {
            console.log(`List ${listName} already exists.`);
        }

        // After ensuring the list exists, retrieve the list details using getByTitle()
        const listDetails = await sp.web.lists.getByTitle(listName).select("Id", "Title")();
        console.log(`List Title: ${listDetails.Title}, List ID: ${listDetails.Id}`);

    }).catch((error: any) => {
        console.error("Error creating or fetching list details:", error);
    });
}


  public async _setFieldTypes(): Promise<void> {
    try {
      const baseUrl = this.context.pageContext.web.absoluteUrl;
      const fields = ['Zeit', 'User', 'Message'];
      for (const field of fields) {
        await SPPost({
          url: `${baseUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
          payload: { "FieldTypeKind": 2, "Title": field, "Required": true }
        });
      }
    } catch (error) {
      console.error("Error setting field types:", error);
    }
  }

  private async _renderListAsync(): Promise<void> {
    try {
      const response = await this._getListData();
      this._renderList(response.value);
    } catch (error) {
      console.error("Error rendering list:", error);
    }
  }

  private _renderList(items: ISPList[]): void {
    let html = '';
    items.forEach((item: ISPList) => {
        const isCurrentUser = item.User === this.context.pageContext.user.displayName;
        html += isCurrentUser ? `
            <div class="${styles.bubble} ${styles.alt}" id="msg_${item.Id}">
              <div class="${styles.txt}">
                <p class="${styles.message}" style="text-align: right;">${item.Message}</p>
                <span class="${styles.timestamp}"><button class="${styles.answer}" id="msg_${item.Id}">Answer</button> ${item.Zeit}</span>
              </div>
            </div>` : `
            <div class="${styles.bubble}">
              <div class="${styles.txt}">
                <p class="${styles.name}">${item.User}</p>
                <p class="${styles.message}">${item.Message}</p>
                <span class="${styles.timestamp}">${item.Zeit}</span>
              </div>
            </div>`;
    });

    const listContainer: Element | null = this.domElement.querySelector('#spListContainer');
    
    if (listContainer) {  // Ensure listContainer is not null
        listContainer.innerHTML = html;
        listContainer.scrollTo(0, listContainer.scrollHeight);
    } else {
        console.error("List container element not found.");
    }

}
    
  

  public _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`,
      SPHttpClient.configurations.v1
    ).then((response: SPHttpClientResponse) => response.json());
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
                PropertyPaneTextField('list', { label: strings.ListFieldLabel }),
                PropertyPaneCheckbox('autofetch', { text: 'Automatically fetch new messages', checked: false }),
                PropertyPaneTextField('interval', { label: strings.IntervalFieldLabel })
              ]
            },
            {
              groupName: strings.StyleGroupName,
              groupFields: [
                PropertyPaneTextField('background', { label: strings.BackgroundFieldLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
