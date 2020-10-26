import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog'; 
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as pnp from 'sp-pnp-js';

export default class CustomDialog extends BaseDialog {
    //public username: string;

    public render(): void { 
        //this.GetUser();

        SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css");
        let loggedUser = "Mickion Mtshali";

        var html:string = "";
        //html +=  '<div class="shadow p-3 mb-5 bg-white rounded">'
        html +=  '<div style="padding: 10px;">';
        html +=  '<h5>Welcome</h5>'; 
        html +=  '<p>You have successfully logged on to Nkusi-it Timesheets Portal.</p>';
        html +=  '<p>If you have any Technical Issues or Questions, please contact your IT Specialist, Mthokozisi Mazibuko on 076 148 6932.</p>';
        //html +=  '<input type="button" id="OkButton"  value="Ok">';  
        html +=  '<button type="button" class="btn btn-success" id="OkButton">Close</button>';
        html +=  '</div>';  
        //html +=  '</div>';
        this.domElement.innerHTML += html;  
        this._setButtonEventHandlers();   
    }

    //Bind event handler to button click
    private _setButtonEventHandlers(): void { 
        const webPart: CustomDialog = this; 
        this.domElement.querySelector('#OkButton').addEventListener('click', () => {    
            //this.paramFromDailog =  document.getElementById("inputParam")["value"] ;   
            this.close();  
        });  
    }

    public getConfig(): IDialogConfiguration {  
        return {  
          isBlocking: false  
        };  
    }  
        
    protected onAfterClose(): void {  
        super.onAfterClose();       
    } 

    private GetUser(){
        pnp.sp.profiles.myProperties.select("AccountName", "PreferredName").get().then(d => {
            alert(JSON.stringify(d));        
        });

        /*pnp.sp.profiles.myProperties.get().then(function(result){
            var userProperties = result.UserProfileProperties;  
            var userPropertyValues = "";  
            userProperties.forEach(function(property) {  
                userPropertyValues += property.Key + " - " + property.Value + "<br/>";  
            });
            alert(userPropertyValues);
        })*/
    }

    
    /*private GetUserProperties(): void {  
        pnp.profiles.myProperties.get().then(function(result) {  
            var userProperties = result.UserProfileProperties;  
            var userPropertyValues = "";  
            userProperties.forEach(function(property) {  
                userPropertyValues += property.Key + " - " + property.Value + "<br/>";  
            });  
            document.getElementById("spUserProfileProperties").innerHTML = userPropertyValues;  
        }).catch(function(error) {  
            console.log("Error: " + error);  
        });  
    }   */
    //Get Current User Display Name
	/*private getSPData(){    
		sp.web.currentUser.get().then((r: CurrentUser) => {
		  return['Title'];
		});
	} */
}
