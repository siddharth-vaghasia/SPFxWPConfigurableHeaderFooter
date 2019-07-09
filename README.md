## SPFx webpart to configure header and footer via SPFx application customizer extension

This repo is a SPFx webpart which allows users to add/modify/delete custom header and footer using SPFx application customizer extension on all modern pages within SP online site. This webpart provides interface to add header and footer so that we don't have to modify code when we need to change header, footer or remove it from site.

Challanges/Drawback with ONLY using SPFx extension for customizing header and footer.
* Header and footer HTML needs to be hardcoded in solution
* Changes to code required if we need to change any text/html in future.
* Redeployment of package and installation
* Diffrent solution would be required for diffrent site collections as we would defintely need diffrent header and footer for each site collection(most of case)
* High maintananence and time consuming for simple task. 

To overcome this drawbacks, this solution comes handy. This is resuable component which can be used by developers to eliminate creating Extension on thier own. Feel free to connect me on [@siddh_me](http://twitter.com/siddh_me/) for any details.

### Features of solution

* Webpart to configure Header and Footer
* Edit functionality if header or footer is already added via this solution.
* Completely Remove header and foooter created via this solution.


### Package and Deploy

Note - If you don't want to build and package on your own, you can directly download package at this [location](./sharepoint/solutions/SPFxAppCustomizerApp.sppkg) and upload to app catalog and install app on required site collection. Skip below steps and directly go to How to use section.

Clone the solution and make sure there is no error before packaging. Try first on local work bench.

Change the pageURL property in /config/serve.json - This should be a valid modern page on your site collection.

```bash
git clone the repo
npm i
gulp serve
```
- Execute the following gulp task to bundle your solution. This executes a release build of your project by using a dynamic label as the host URL for your assets. This URL is automatically updated based on your tenant CDN settings:
```bash
gulp bundle --ship
```
- Execute the following task to package your solution. This creates an updated webpart.sppkg package on the sharepoint/solution folder.
```bash
gulp package-solution --ship
```
- Upload or drag and drop the newly created client-side solution package to the app catalog in your tenant.
- Based on your tenant settings, if you would not have CDN enabled in your tenant, and the includeClientSideAssets setting would be true in the package-solution.json, the loading URL for the assets would be dynamically updated and pointing directly to the ClientSideAssets folder located in the app catalog site collection.


### How to Use Solution

For detailed instruction refer this [Document](./How%20to%20Configure%20Solution.docx)
* Once app is deployed to app catalog sucessfully.
* Install app to required site collection
* Create new modern page. Add 'WpConfigureApplicationCustomizerWebPart' webpart to page. Publish the page.
* Enter either or both Header HTML and Footer HTML and click on Register Custom Action button
* On Success message - Refresh the page and you would see your header and footer on all your modern pages.
* To Edit/Remove, go to same page again and Use 'Update Custom Action' or 'Remove Custom Action' button.

### High level design of Solution

* SPFx solution with 2 components 1. SPFx WebPArt 2. SPFx Extension Application Customizer
* No Javascript framework based solution to keep it light
* Disables Automatic activation of SPFx extension when app is installed.
* Register Custom action with ClientSideComponentId of Extension component
* Passes parameters to Extension with ClientSideComponentProperties
* SPFx webpart uses Bootstrap UI framework and Jquery

For any issue or help, Buzz me on [@siddh_me](http://twitter.com/siddh_me/)

> Sharing is caring!