
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse,   
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import styles from './WpConfigureApplicationCustomizerWebPart.module.scss';
import * as strings from 'WpConfigureApplicationCustomizerWebPartStrings';
import * as jQuery from 'jquery';
import * as bootstrap from 'bootstrap';
import { SPComponentLoader } from '@microsoft/sp-loader';
//require('bootstrap');
export interface IWpConfigureApplicationCustomizerWebPartProps {
  description: string;
}

export default class WpConfigureApplicationCustomizerWebPart extends BaseClientSideWebPart<IWpConfigureApplicationCustomizerWebPartProps> {
  
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.wpConfigureApplicationCustomizer }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">SPFx Application Customizer - Configurable Header and Footer</span>
              <p class="${ styles.subTitle }">This webpart can be used to add customized header and footer via SPFx extension application customizer.</p>
              <p class="${ styles.description }">Add your HTML for customized Header and Footer on SharePoint Online Site. Application Customizers provide access to well-known locations on SharePoint pages that you can modify based on your business and functional requirements. For example, you can create dynamic header and footer experiences that render across all the pages in SharePoint Online.</p>
              <a target="_blank" href="https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/using-page-placeholder-with-extensions" class="${ styles.button }">
                <span class="${ styles.label }">More on Application Customizer</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;

      let cssURL = "https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css";
      SPComponentLoader.loadCss(cssURL);
      
      SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css");
      SPComponentLoader.loadScript("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js");
      this.domElement.innerHTML += `<div id="maincontent"></br> </br>
      
      <div id="existMessage" style="display:none" class="alert alert-info">
  <strong>Info!</strong> We found you already have custom header and footer added. Feel free to Edit or Remove Customization.
</div>

      <div class="form-group">
      <label for="Enter HTML to be added in ">Header HTML</label> 
      <textarea id="headerText" name="headerText" cols="40" rows="5" aria-describedby="Enter HTML to be added in HelpBlock" class="form-control"></textarea> 
      <span id="Enter HTML to be added in HelpBlock" class="form-text text-muted">Text in this control will be added in Top placeholder this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top)</span>
    </div>
    <div class="form-group">
      <label for="textarea">Footer HTML</label> 
      <textarea id="footerText" name="footerText" cols="40" rows="5" aria-describedby="textareaHelpBlock" class="form-control"></textarea> 
      <span id="textareaHelpBlock" class="form-text text-muted">Text in this control will be added in Bottom placeholder this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom)</span>
    </div> 

    <div style="display:none">
   
    <input id="currentActionId" ></input> 
   
    </div> 
   
    
    <div class="form-group row">
      <div class="offset-4 col-8">
        <button name="btnRegister" type="submit" class="btn btn-primary" id="btnRegister">Register Custom Action</button>
        <button name="btnRemove" style="display:none" type="submit" class="btn btn-secondary" id="btnRemove">Remove Custom Action</button>
      </div>
    </div> 
    </div>
    <div class="form-group row">
    </br>
    </br>
    <div style="display:none" id="successmessage" class="alert alert-success">
    <strong>Success!</strong> Registered Custom action successfully, refresh page to see customized header and footer.
  </div>

    `;

    
      //this.domElement.innerHTML += '<textarea id ="headerText" rows = "5" cols = "50" name = "description">Enter Header HTML</textarea>';
      //this.domElement.innerHTML += '<button type="button" id="btnRegister">Click Me!</button>';
      //this.domElement.innerHTML += '<button type="button" id="btnRemove">Remove Me!</button>';

      this._setButtonEventHandlers();
  }

  private _setButtonEventHandlers(): void {
    const webPart: WpConfigureApplicationCustomizerWebPart = this;
    console.log(jQuery("#btnRegister").text());
    this.domElement.querySelector('#btnRegister').addEventListener('click', () => {
     
      //this.getData();
     
      this.setCustomAction();
      
    });

    this.domElement.querySelector('#btnRemove').addEventListener('click', () => {
     
      //this.getData();
     
      this.removeCustomAction();
    });

    this.getCustomAction();
    
 }

 /*private addUCAforApplicationCustomizer(){
alert('addUCAforApplicationCustomizer')
  var clientContext = SP.ClientContext.get_current();
var userCustomActions = clientContext.get_site().get_userCustomActions();
clientContext.load(userCustomActions);

var action = userCustomActions.add();

action.set_location("ClientSideExtension.ApplicationCustomizer");
action.set_title("SettingHeaderFooterwithAppCustomizer");
action.set_description("This user action is of type  application customizer to add header footer or custom javascript via SFPx extension");
action.set_clientsidecomponentid('fc14316e-72db-4860-8283-458154eaef3e');
action.set_clientsidecomponentproperties("{\"testMessage\":\"From my\"}");


action.update();

clientContext.load(action);
clientContext.executeQueryAsync(function(){
    alert("success");
},function(){
    alert("error");
});
}*/


private _getdigest(): Promise<any> {
 

  const spOpts: ISPHttpClientOptions = {
  };

  return this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + `/_api/contextinfo`, SPHttpClient.configurations.v1,spOpts)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
}

protected removeCustomAction(){
  try {
    var title = 'SettingHeaderFooterwithAppCustomizer';
    var description = 'This user action is of type  application customizer to add header footer or custom javascript via SFPx extension';
      this._getdigest()
        .then((digrestJson) => {
          console.log(digrestJson);

          const digest = digrestJson.FormDigestValue;

        
          const headers = {
              'X-RequestDigest': digest,
              "content-type": "application/json;odata=verbose",
          };
          const spOpts: ISPHttpClientOptions = {
            
          };        
                  //Remove application customizer
                  this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "/_api/site/UserCustomActions(@v0)/deleteObject()?@v0=guid'" +  jQuery("#currentActionId").val() + "'", SPHttpClient.configurations.v1,spOpts)
                    .then((response: SPHttpClientResponse) => {
                      console.log( response);
                      jQuery("#maincontent").hide();
                      jQuery("#successmessage").html("<strong>Success!</strong> Removed Custom Action Successfully. Refresh the page to view updated header and footer")
                      jQuery("#successmessage").show();
                         
                    });
                  
              
      });
  } catch (error) {
      console.error(error);
  }
}



protected getCustomAction() {
  try {
    var title = 'SettingHeaderFooterwithAppCustomizer';
    var description = 'This user action is of type  application customizer to add header footer or custom javascript via SFPx extension';
      this._getdigest()
        .then((digrestJson) => {
          console.log(digrestJson);

          const digest = digrestJson.FormDigestValue;

        
          const headers = {
              'X-RequestDigest': digest,
              "content-type": "application/json;odata=verbose",
          };
          const spOpts: ISPHttpClientOptions = {
            
          };

          this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/site/UserCustomActions`, SPHttpClient.configurations.v1,spOpts)
          .then((response: SPHttpClientResponse) => {
             
             response.json().then((responseJSON: any) => {  
              console.log(responseJSON);  

              responseJSON.value.forEach(element => {
                  if(element.Title == title)
                  {
                    console.log(element);
                    if(JSON.parse(element.ClientSideComponentProperties).Top != "" || JSON.parse(element.ClientSideComponentProperties).Bottom !="")
                    {
                     
                     jQuery("#existMessage").show();
                     jQuery("#btnRegister").text("Update Custom Action");
                     jQuery("#btnRemove").show();
                     jQuery("#headerText").val(JSON.parse(element.ClientSideComponentProperties).Top) ;
                     jQuery("#footerText").val(JSON.parse(element.ClientSideComponentProperties).Bottom);
                     
                    }
                  }
                  jQuery("#currentActionId").val(element.Id);
              });
            });  
           });  
      });
  } catch (error) {
      console.error(error);
  }
}

protected setCustomAction() {
  try {
    var title = 'SettingHeaderFooterwithAppCustomizer';
    var description = 'This user action is of type  application customizer to add header footer or custom javascript via SFPx extension';

    var headtext =  document.getElementById("headerText")["value"] ;
    var foottext =  document.getElementById("footerText")["value"] ;
    
    
      this._getdigest()
        .then((digrestJson) => {
          console.log(digrestJson);

          const digest = digrestJson.FormDigestValue;

          const payload = JSON.stringify({
             
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Title: title,
              Description: description,
              ClientSideComponentId: 'fc14316e-72db-4860-8283-458154eaef3e',
              ClientSideComponentProperties: JSON.stringify({Top:headtext,Bottom:foottext }),
          });
          const headers = {
              'X-RequestDigest': digest,
              "content-type": "application/json;odata=verbose",
          };
          const spOpts: ISPHttpClientOptions = {
            body:payload
          };

          //Remove application customizer
          this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + "/_api/site/UserCustomActions(@v0)/deleteObject()?@v0=guid'" +  jQuery("#currentActionId").val() + "'", SPHttpClient.configurations.v1,spOpts)
          .then((response: SPHttpClientResponse) => {
                // Register a new one...
                this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl + `/_api/site/UserCustomActions`, SPHttpClient.configurations.v1,spOpts)
                .then((response: SPHttpClientResponse) => {
                  console.log( response.json());
                  jQuery("#maincontent").hide();
                  jQuery("#successmessage").html("<strong>Success!</strong> Custom Action Registered Successfully. Refresh the page to view updated header and footer")
                  jQuery("#successmessage").show();
                });
                });
               
          });

  } catch (error) {
      console.error(error);
  }
}

  protected getData(){
    
    this._getListData()
        .then((response) => {
          console.log(response);
        });
  }

  private _getListData(): Promise<any> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
