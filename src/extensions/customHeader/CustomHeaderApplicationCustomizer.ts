import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,  
  PlaceholderName,  
  PlaceholderProvider  
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CustomHeaderApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CustomHeaderApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomHeaderApplicationCustomizerProperties {
  // This is an example; replace with your own property
 
  Top: string;
  Bottom: string;
  
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CustomHeaderApplicationCustomizer
  extends BaseApplicationCustomizer<ICustomHeaderApplicationCustomizerProperties> {
    private _topPlaceholder: PlaceholderContent | undefined;
    private _bottomPlaceholder: PlaceholderContent | undefined;
  @override
  public onInit(): Promise<void> {
  /*  Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let topmessage: string = this.properties.Top;
    if (!topmessage) {
      topmessage = '(No properties were provided.)';
       //CODE to add custom html at header
      let topPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);  
      if (topPlaceholder) {  
        topPlaceholder.domElement.innerHTML = '<div><div style="text-align:center" > This is to demo SPFx extension to customize app header.'  +  this.properties.myprop +'</div> </div>';

      }
      
      //CODE to add custom html at footer
            let bottomPlaceholder: PlaceholderContent = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);  
            if (bottomPlaceholder) {  
              bottomPlaceholder.domElement.innerHTML = '<div style="background-color: red;"><div style="text-align:center;" > This is to demo SPFx extension to customize app footer. </div> </div>';       
            }
      */

      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

      // Wait for the placeholders to be created (or handle them being changed) and then
	// render.
      this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
	
      return Promise.resolve<void>();
    }
      

  private _renderPlaceHolders(): void {

    console.log("HelloWorldApplicationCustomizer._renderPlaceHolders()");
 		console.log(
 			"Available placeholders: ",
 			this.context.placeholderProvider.placeholderNames
 				.map(name => PlaceholderName[name])
 				.join(", ")
 		);

 		// Handling the top placeholder
 		if (!this._topPlaceholder) {
 			this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
 				PlaceholderName.Top,
 				{ onDispose: this._onDispose }
 			);

 			// The extension should not assume that the expected placeholder is available.
 			if (!this._topPlaceholder) {
 				console.error("The expected placeholder (Top) was not found.");
 				return;
 			}

 			if (this.properties) {
 				let topString: string = this.properties.Top;
 				if (!topString) {
 					topString = "(Top property was not defined.)";
 				}

 				if (this._topPlaceholder.domElement) {
 					this._topPlaceholder.domElement.innerHTML = this.properties.Top;
 				}
 			}
 		}

 		// Handling the bottom placeholder
 		if (!this._bottomPlaceholder) {
 			this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
 				PlaceholderName.Bottom,
 				{ onDispose: this._onDispose }
 			);

 			// The extension should not assume that the expected placeholder is available.
 			if (!this._bottomPlaceholder) {
 				console.error("The expected placeholder (Bottom) was not found.");
 				return;
 			}

 			if (this.properties) {
 				let bottomString: string = this.properties.Bottom;
 				if (!bottomString) {
 					bottomString = "(Bottom property was not defined.)";
 				}

 				if (this._bottomPlaceholder.domElement) {
 					this._bottomPlaceholder.domElement.innerHTML = this.properties.Bottom;
 				}
 			}
     }
  }
  private _onDispose(): void {
    console.log('[CustomHeaderApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
  
