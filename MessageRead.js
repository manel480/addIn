var myMSALObj;
var loginRequest = { scopes: ["openid", "profile", "User.Read", "Calendars.ReadWrite", "Place.ReadWrite.All"] };

(function () {
    "use strict";

    Office.onReady(function () {
        $(document).ready(function () {
            startMsal();


            Office.context.roamingSettings.remove("entorno");
            Office.context.roamingSettings.saveAsync();


            document.getElementById("roomSelected").onclick = seeRoomSelected;
            document.getElementById("sendEntorno").onclick = saveEntorno;
            
        });
    });

    function startMsal() {
        myMSALObj = new Msal.UserAgentApplication(
            {
                auth: {
                    clientId: "2da12d3f-b8a4-402c-a335-37714c452408",
                    authority: "https://login.microsoftonline.com/6022b001-3112-4886-87c4-bfcbaebe61a2",
                    redirectUri: "https://localhost:44328/MessageRead.html"
                },
                cache: {
                    cacheLocation: "sessionStorage",
                    storeAuthStateInCookie: false,
                }
            }
        );
        if (!myMSALObj.getAccount()) {
            myMSALObj.loginPopup(loginRequest)
                .then(loginResponse => {
                    if (myMSALObj.getAccount()) {
                        showWelcomeMessage(myMSALObj.getAccount());
                        
                    }
                }).catch(error => {
                    console.log(error);
                });
        }
        else {
            showWelcomeMessage(myMSALObj.getAccount());
          
        }
    }

    function showWelcomeMessage(account) {
        var settingValue = Office.context.roamingSettings.get("entorno");
        if (settingValue !== undefined) {
            document.getElementById("card-div").classList.remove('d-none');
            document.getElementById("card-div2").classList.remove('d-none');
            document.getElementById("card-div3").classList.add('d-none');
            document.getElementById("welcomeMessage").innerHTML = `Hola ${account.name} <br/> ` + settingValue;
        }
        else {
            document.getElementById("card-div").classList.remove('d-none');
            document.getElementById("card-div3").classList.remove('d-none');
            document.getElementById("welcomeMessage").innerHTML = `Hola ${account.name}`;
        }
    }

    function seeRoomSelected() {
        const idRoom = document.getElementById("selectRoom").value;
        if (idRoom != "") {
            const finalUrl = "https://graph.microsoft.com/v1.0/places/microsoft.graph.room/" + idRoom;
            if (myMSALObj.getAccount()) {
                getTokenPopup()
                    .then(response => {
                        callMSGraph(finalUrl, response.accessToken);
                    }).catch(error => {
                        console.log(error);
                    });
            }
        }
    }

    function getTokenPopup() {
        return myMSALObj.acquireTokenSilent(loginRequest)
            .catch(error => {
                console.log(error);
                return myMSALObj.acquireTokenPopup(loginRequest)
                    .then(tokenResponse => {
                        return tokenResponse;
                    }).catch(error => {
                        console.log("Hay error" + error);
                    });
            });
    }


    function callMSGraph(endpoint, token) {
        const headers = new Headers();
        const bearer = `Bearer ${token}`;
        headers.append("Authorization", bearer);
        const options = { method: "GET", headers: headers };

        fetch(endpoint, options)
            .then(response => response.json())
            .then(response => {
                //document.getElementById("salas").innerHTML = "Sala selececcionada: " + response.emailAddress;

                if (response.hasOwnProperty("error"))
                    console.log("La sala no existe");
                else {
                    const locations = [{
                        id: response.emailAddress,
                        type: Office.MailboxEnums.LocationType.Room
                    }];
                    Office.context.mailbox.item.enhancedLocation.addAsync(locations, (result) => {
                        if (!result.status === Office.AsyncResultStatus.Succeeded) {
                            console.error(`Room no exist.`);
                        }
                    });
                }
            })
            .catch(error => console.log(error))
    }

    function saveEntorno() {
        var entorno = document.getElementById("valueEntorno").value;
        Office.context.roamingSettings.set("entorno", entorno);
	    Office.context.roamingSettings.saveAsync(function (result) {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
                console.error(`Action failed `);
            } else {
                var settingValue = Office.context.roamingSettings.get("entorno");
                document.getElementById("muestraRoamingValue").innerHTML = "El valor de la variable ENTORNO es: " +settingValue;
                console.log(`Settings saved `);
            }
        });
        showWelcomeMessage(myMSALObj.getAccount());
    }
})();