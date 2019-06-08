## SPFx webpart to configure header and footer via SPFx application customizer extension

This repo is a SPFx webpart which allows users to add/modify/delete custom header and footer using SPFx application customizer extension all modern pages within SP online site. This webpart provide interface to add header and footer so that we don't have to modify code when we need to change header, footer or remove it from site.

Challanges/Drawback with ONLY using SPFx extension for customizing header and footer.
* Header and footer HTML needs to be hardcoded in solution
* Changes to code required if we need to change any text/html in future.
* Redeployment of package and installation
* Diffrent solution would be required for diffrent site collections as we would defintely need diffrent header and footer for each site collection(most of case)
* High maintananence and time consuming for simple task. 

To overcome this drawbacks, this solution comes handy. Feel free to connect on siddh.vaghasia@gmail.com for any details.

### Features of solution

* WebPart to configure Header and Footer
* Edit functionality if header or footer is already added via this solution
* Completely Remove header and foooter created via this solution

### How to Use Solution
For detailed instruction refer this [Document]  (https://github.com/siddharth-vaghasia/SPFxWPConfigurableHeaderFooter/blob/master/How%20to%20Configure%20Solution.docx "How to configure solution")
* Download code
* Test your webpart from local server (to confirm it is working)
* Enable Office 365 CDN [Host webpart from Office 365 CDN](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/hosting-webpart-from-office-365-cdn "Office 365 CDN")
* Package solution, Upload to app catalog
* Install app to required site collection
* Create new modern page. Add 'WpConfigureApplicationCustomizerWebPart' webpart to page. Publish the page.
* Enter either or both Header HTML and Footer HTML and click on Register Custom Action button
* On Success message - Refresh the page and you would see your header and footer on all your modern pages.
* To Edit/Remove, go to same page again and Use 'Update Custom Action' or 'Remove Custom Action' button.

### High level design of Solution

* SPFx solution with 2 components 1. SPFx WebPArt 2. SPFx Extension Application Customizer
* Disables Automatic activation of SPFx extension when app is installed.
* Register Custom action with ClientSideComponentId of Extension component
* Passes parameters to Extension with ClientSideComponentProperties
* SPFx webpart uses Bootstrap UI framework and Jquery

For any issue or help, contact me on siddh.vaghasia@gmail.com