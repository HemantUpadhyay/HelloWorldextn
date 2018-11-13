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
    this.Initiate()
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }
    //Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);
    //Dialog.alert("after Initiate call.")
    return Promise.resolve();
  }

  private _getNewPageStatus():void {
    let responseText: string = "";
    var functionSIteIDUrl = "https://functesthelloworld.azurewebsites.net/api/HttpTrigger1?code=zWFUcRwMIeXtUaCCYP8BOWYFa5jQn5SAE9/hHqFL/6Uk/mfavUhw0Q==&name=Hemant";
    this.context.spHttpClient.get(functionSIteIDUrl,SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      if(response.ok)
          {
            response.json().then((responseJSON: JSON) => {
                //._callPiwikScript(responseJSON.value);
                responseText = JSON.stringify(responseJSON);
                if (response.ok) {
                    //resultMsg.style.color = "green";
                    Dialog.alert("done");
                    console.log(response);
                } else {
                    //resultMsg.style.color = "red";
                    Dialog.alert("fail");
                    console.log(response);
                }
              });
          }
      // response.json().then((responseJSON: JSON) => {
      //   responseText = JSON.stringify(responseJSON);
      //       if (response.ok) {
      //           //resultMsg.style.color = "green";
      //           Dialog.alert("done");
      //           console.log(response);
      //       } else {
      //           //resultMsg.style.color = "red";
      //           Dialog.alert("fail");
      //           console.log(response);
      //       }
      // })
           
  }).catch((e) => {
    console.log(e);
  });
}



  private async getNewPageStatus() {
    Dialog.alert("Inside getNewPageStatus function")
    const currentWebUrl = this.context.pageContext.web.absoluteUrl;
    //const pageName = 'DynamicPage.aspx'
    var functionSIteIDUrl = "https://functesthelloworld.azurewebsites.net/api/HttpTrigger1?code=zWFUcRwMIeXtUaCCYP8BOWYFa5jQn5SAE9/hHqFL/6Uk/mfavUhw0Q==&name=Hemant";

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/jsonp");
    requestHeaders.append("Cache-Control", "private"); 

    const postOptions : RequestInit = {
    headers: requestHeaders,
    //body: `{\r\n    siteURL: '${currentWebUrl}',\r\n    pageName: '${pageName}' \r\n}`,
    body: `{\r\n    siteURL: '${currentWebUrl}'\r\n}`,
    method: "GET"
    };

    let responseText: string = "";
    let createPageStatus: string = "";
    console.log("Wait started for Creating page");    
    $.ajax({
      //JSONP API       
      url: "https://functesthelloworld.azurewebsites.net/api/HttpTrigger1?code=zWFUcRwMIeXtUaCCYP8BOWYFa5jQn5SAE9/hHqFL/6Uk/mfavUhw0Q==&name=Hemant",
      method: "GET",
      dataType: "jsonp",
      async:false,
      headers: {'Content-Type':'application/json','Cache-Control':'private','Access-Control-Allow-Origin':'*'},
      success: function (data) {
        //alert('success');
        console.log(data);
      },
      error: function (jqXHR, textStatus,error) {
        console.log('error...!!!'+error)
      }
    });
    // await fetch(functionSIteIDUrl, postOptions).then((response) => {
    //     console.log("Response returned");
    //     if (response.ok) {
    //       Dialog.alert(`Page status returned`);
    //       return response.json()          
    //     }
    //     else
    //     {
    //         var errMsg = "Error detected while adding site page. Server response wasn't OK ";
    //         console.log(errMsg);
    //     } 
    //   }).then((responseJSON: JSON) => {
    //     responseText = JSON.stringify(responseJSON).trim();
    //     console.log(responseText);
    //     if(responseText.toLowerCase().indexOf("success") > 0)
    //         {
    //           console.log("success feedback");
    //           //to make another call for next azure method on success of 1st method
    //           this.insertWebPartToPage();
    //         }
    //     if(responseText.toLowerCase().indexOf("error") > 0)
    //         {
    //           console.log("web call errored");
    //         }
    //   }
    // ).catch ((response: any) => {
    //   let errMsg: string = `WARNING - error when calling URL ${functionSIteIDUrl}. Error = ${response.message}`;
    //   console.log(errMsg);
    // });



    console.log("wait finished");
  }

  private async insertWebPartToPage()
  {
    const currentWebUrl: string = this.context.pageContext.web.absoluteUrl;
    //const pageName = 'DynamicPage.aspx'    
    var functionInsertWebPartUrl : string = "https://functesthelloworld.azurewebsites.net/api/HttpTrigger1?code=zWFUcRwMIeXtUaCCYP8BOWYFa5jQn5SAE9/hHqFL/6Uk/mfavUhw0Q==&name=Hemant";
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "private"); 
    //requestHeaders.append("Data", "jsonp"); 

    const getOption : RequestInit = {
    headers: requestHeaders,
    //body: `{\r\n    siteURL: '${currentWebUrl}',\r\n    pageName: '${pageName}' \r\n}`,
    body: `{\r\n    siteURL: '${currentWebUrl}'\r\n}`,
    method: "GET"
    };

    let responseText: string = "";
    let createPageStatus: string = "";
    console.log("Wait started for adding Web Part");
    await fetch(functionInsertWebPartUrl, getOption).then((response) => {
        console.log("Response returned");
        if (response.ok) {
          Dialog.alert(`Insert web part alert`);
          return response.json()
        }
        else
        {
            var errMsg = "Error detected while adding web-part to site page. Server response wasn't OK ";
            console.log(errMsg);
        } 
      }).then((responseJSON: JSON) => {
        responseText = JSON.stringify(responseJSON).trim();
        console.log(responseText);
        if(responseText.toLowerCase().indexOf("success") > 0)
            {
              console.log("Web-part add success");
            }
        if(responseText.toLowerCase().indexOf("error") > 0)
            {
              console.log("web call errored");
            }
      }
    ).catch ((response: any) => {
      let errMsg: string = `WARNING - error when calling URL ${functionInsertWebPartUrl}. Error = ${response.message}`;
      console.log(errMsg);
    });
    console.log("wait for web Part add finished");
  }
  
  private async Initiate() {    
    //await this.getNewPageStatus();
    this._getNewPageStatus();
  }
}
