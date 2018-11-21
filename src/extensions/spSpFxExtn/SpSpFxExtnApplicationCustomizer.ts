import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import {
  SPHttpClient,
  SPHttpClientResponse ,
  SPHttpClientConfiguration   
 } from '@microsoft/sp-http';
import * as strings from 'SpSpFxExtnApplicationCustomizerStrings';
import * as $ from 'jquery'
const LOG_SOURCE: string = 'SpSpFxExtnApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpSpFxExtnApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpSpFxExtnApplicationCustomizer
  extends BaseApplicationCustomizer<ISpSpFxExtnApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    //Dialog.alert("before Initiate call.");
    this._Initiate()
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
	
    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }
    return Promise.resolve();
  }

  private _GetNewPageStatus():void {
    //let responseText: string = "";
    console.log("Wait started for Creating page"); 
    var functionSIteIDUrl = "https://functesthelloworld.azurewebsites.net/api/HttpTrigger1?code=zWFUcRwMIeXtUaCCYP8BOWYFa5jQn5SAE9/hHqFL/6Uk/mfavUhw0Q==&name=Hemant";
    this.context.spHttpClient.get(functionSIteIDUrl,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if(response.ok)
      {
        response.json().then((responseJSON: JSON) => {
            //._callPiwikScript(responseJSON.value);
            // = JSON.stringify(responseJSON);
            if (response.ok) {
                //resultMsg.style.color = "green";
                Dialog.alert("page status check done");                
                console.log(response);
                this._InsertWebPartToPage();
            } else {
                //resultMsg.style.color = "red";
                Dialog.alert("fail");
                console.log(response);
            }
          });
      }     
    }).catch((e) => {
      console.log(e);
    });
  }

  private async _InsertWebPartToPage()
  {
    console.log("Wait started");     
    //let responseText: string = "";
    var functionSIteIDUrl = "https://functesthelloworld.azurewebsites.net/api/HttpTrigger1?code=zWFUcRwMIeXtUaCCYP8BOWYFa5jQn5SAE9/hHqFL/6Uk/mfavUhw0Q==&name=Hemant";
    this.context.spHttpClient.get(functionSIteIDUrl,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if(response.ok)
      {
        response.json().then((responseJSON: JSON) => {
            //._callPiwikScript(responseJSON.value);
            // = JSON.stringify(responseJSON);
            if (response.ok) {
                //resultMsg.style.color = "green";
                Dialog.alert("WEP part add done");
                console.log(response);
            } else {
                //resultMsg.style.color = "red";
                Dialog.alert("fail");
                console.log(response);
            }
          });
      }     
    }).catch((e) => {
      console.log(e);
    });
    console.log("wait for web Part add finished");
  }
  
  private async _Initiate() {    
    //await this.getNewPageStatus();
    this._GetNewPageStatus();
  }
}
