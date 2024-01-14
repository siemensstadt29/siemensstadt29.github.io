$.ajax({
    url: "https://login.microsoftonline.com/d2612661-ecdd-4ae3-af61-2ba8db109ab7/oauth2/v2.0/token",
    headers: {
        "Host": "login.microsoftonline.com",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8"
    },
    method: "POST",
    data: {
        "client_id": "971008be-7612-4768-be87-450cfad64814",
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": "0b48Q~r2jQj-KMap.gABvbHKyr1PPcuNTqFKwc0N",
        "grant_type": "client_credentials"
    },
    success: function(result) {
        console.log(JSON.stringify(result));
    }
});