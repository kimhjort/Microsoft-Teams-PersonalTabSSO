let accessToken;

$(document).ready(function () {
    microsoftTeams.initialize();

    getClientSideToken()
        .then((clientSideToken) => {
            console.log("clientSideToken: " + clientSideToken);
            return getServerSideToken(clientSideToken);
        })
        .catch((error) => {
            console.log(error);
            if (error === "invalid_grant") {
                // Display in-line button so user can consent
                $("#divError").text("Error while exchanging for Server token - invalid_grant - User or admin consent is required.");
                $("#divError").show();
                $("#consent").show();
                $("#errorPanel").show();
            } else {
                // Something else went wrong
            }
        });
});

function requestConsent() {
    getToken()
        .then(data => {
            $("#consent").hide();
            $("#divError").hide();
            $("#errorPanel").hide();
            accessToken = data.accessToken;
            microsoftTeams.getContext((context) => {
                getUserInfo(context.userPrincipalName);
                getPhotoAsync(accessToken);
                getManagerInfo();
                getMyPresence();
            });
        });
}

function getToken() {
    return new Promise((resolve, reject) => {
        microsoftTeams.authentication.authenticate({
            url: window.location.origin + "/Auth/Start",
            width: 600,
            height: 535,
            successCallback: result => {
                resolve(result);
            },
            failureCallback: reason => {
                reject(reason);
            }
        });
    });
}

function getClientSideToken() {
    return new Promise((resolve, reject) => {
        microsoftTeams.authentication.getAuthToken({
            successCallback: (result) => {
                resolve(result);
            },
            failureCallback: function (error) {
                reject("Error getting token: " + error);
            }
        });
    });
}

function getServerSideToken(clientSideToken) {
    return new Promise((resolve, reject) => {
        microsoftTeams.getContext((context) => {
            var scopes = ["https://graph.microsoft.com/User.Read"];
            fetch('/GetUserAccessToken', {
                method: 'get',
                headers: {
                    "Content-Type": "application/text",
                    "Authorization": "Bearer " + clientSideToken
                },
                cache: 'default'
            })
                .then((response) => {
                    if (response.ok) {
                        return response.text();
                    } else {
                        reject(response.error);
                    }
                })
                .then((responseJson) => {
                    if (IsValidJSONString(responseJson)) {
                        if (JSON.parse(responseJson).error)
                            reject(JSON.parse(responseJson).error);
                    } else if (responseJson) {
                        accessToken = responseJson;
                        console.log("Exchanged token: " + accessToken);
                        getUserInfo(context.userPrincipalName);
                        getPhotoAsync();
                        getManagerInfo();
                        getMyPresence();
                    }
                });
        });
    });
}

function IsValidJSONString(str) {
    try {
        JSON.parse(str);
    } catch (e) {
        return false;
    }
    return true;
}

function getUserInfo(principalName) {
    if (principalName) {
        let graphUrl = "https://graph.microsoft.com/v1.0/users/" + principalName;
        $.ajax({
            url: graphUrl,
            type: "GET",
            beforeSend: function (request) {
                request.setRequestHeader("Authorization", `Bearer ${accessToken}`);
            },
            success: function (profile) {
                let profileDiv = $("#divGraphProfile");
                profileDiv.empty();
                for (let key in profile) {
                    if ((key[0] !== "@") && profile[key]) {
                        $("<div>")
                            .append($("<b>").text(key + ": "))
                            .append($("<span>").text(profile[key]))
                            .appendTo(profileDiv);
                    }
                }

                $("#yourName").text(profile.displayName);
                $("#divGraphProfile").show();
            },
            error: function (xhr, status, errorThrown) {
                console.log("getManagerInfo Failed: " + errorThrown);
            },
            complete: function (data) {
            }
        });
    }
}

//https://graph.microsoft.com/v1.0/me/presence //my presence
//https://graph.microsoft.com/v1.0/users/d4957c9d-869e-4364-830c-d0c95be72738/presence //gets a users presence

function getMyPresence() {
    let graphUrl = "https://graph.microsoft.com/v1.0/me/presence";
    $.ajax({
        url: graphUrl,
        type: "GET",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", `Bearer ${accessToken}`);
        },
        success: function (presencestatus) {
            if (presencestatus) {
                $("#myPresence").removeClass("unknown").text("").attr('title', presencestatus.availability).tooltip();
                $("#myPresence").addClass(presencestatus.availability.toLowerCase());

                if (presencestatus.availability.toLowerCase() === "donotdisturb") {
                    $("#myPresence").text("➖");
                } else if (presencestatus.availability.toLowerCase() === "available") {
                    $("#myPresence").text("✔");
                } else if (presencestatus.availability.toLowerCase() === "away") {
                    $("#myPresence").text("🕡");
                }
            }
        },
        error: function (xhr, status, errorThrown) {
            console.log("getMyPresence Failed: " + errorThrown);
        },
        complete: function (data) {
        }
    });
}

function getUsersPresence(UserId) {
    let graphUrl = `https://graph.microsoft.com/v1.0/users/${UserId}/presence`;
    $.ajax({
        url: graphUrl,
        type: "GET",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", `Bearer ${accessToken}`);
        },
        success: function (presencestatus) {
            if (presencestatus) {
                $("#managerPresence").removeClass("unknown").text("").attr('title', presencestatus.availability).tooltip();
                $("#managerPresence").addClass(presencestatus.availability.toLowerCase());

                if (presencestatus.availability.toLowerCase() === "donotdisturb") {
                    $("#managerPresence").text("➖");
                } else if (presencestatus.availability.toLowerCase() === "available") {
                    $("#managerPresence").text("✔");
                } else if (presencestatus.availability.toLowerCase() === "away") {
                    $("#managerPresence").text("🕡");
                }
            }
        },
        error: function (xhr, status, errorThrown) {
            console.log("getUsersPresence Failed: " + errorThrown);
        },
        complete: function (data) {
        }
    });
}

function getManagerInfo() {
    let graphUrl = "https://graph.microsoft.com/v1.0/me/manager";
    $.ajax({
        url: graphUrl,
        type: "GET",
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", `Bearer ${accessToken}`);
        },
        success: function (profile) {
            let profileDiv = $("#divManagerProfile");
            profileDiv.empty();
            for (let key in profile) {
                if ((key[0] !== "@") && profile[key]) {
                    $("<div>")
                        .append($("<b>").text(key + ": "))
                        .append($("<span>").text(profile[key]))
                        .appendTo(profileDiv);
                }
            }

            $("#divManagerProfile").show();
            $("#managerName").text(profile.displayName);

            // Gets manager image
            getManagerPhotoAsync(profile.id);
            getUsersPresence(profile.id);
        },
        error: function (xhr, status, errorThrown) {
            console.log("getManagerInfo Failed: " + errorThrown);
        },
        complete: function (data) {
        }
    });
}

// Gets the user's photo
function getPhotoAsync() {
    let graphPhotoEndpoint = 'https://graph.microsoft.com/v1.0/me/photos/240x240/$value';
    let request = new XMLHttpRequest();
    request.open("GET", graphPhotoEndpoint, true);
    request.setRequestHeader("Authorization", `Bearer ${accessToken}`);
    request.setRequestHeader("Content-Type", "image/png");
    request.responseType = "blob";
    request.onload = function (oEvent) {
        let imageBlob = request.response;
        if (imageBlob) {
            let urlCreater = window.URL || window.webkitURL;
            let imgUrl = urlCreater.createObjectURL(imageBlob);
            $("#userPhoto").attr('src', imgUrl);
            $("#userPhoto").show();
        }
    };
    request.send();
}

// Gets the user's photo
function getManagerPhotoAsync(managerId) {
    let graphPhotoEndpoint = 'https://graph.microsoft.com/v1.0/users/' + managerId + '/photos/240x240/$value';
    let request = new XMLHttpRequest();
    request.open("GET", graphPhotoEndpoint, true);
    request.setRequestHeader("Authorization", `Bearer ${accessToken}`);
    request.setRequestHeader("Content-Type", "image/png");
    request.responseType = "blob";
    request.onload = function (oEvent) {
        let imageBlob = request.response;
        if (imageBlob) {
            let urlCreater = window.URL || window.webkitURL;
            let imgUrl = urlCreater.createObjectURL(imageBlob);
            $("#ManagerPhoto").attr('src', imgUrl);
            $("#ManagerPhoto").show();
        }
    };
    request.send();
}