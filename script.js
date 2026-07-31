const WEB_APP_URL =
"https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxx/exec";

async function askAI(){

const question =
document.getElementById("question").value;

const response = await fetch(WEB_APP_URL,{

method:"POST",

headers:{
"Content-Type":"application/json"
},

body:JSON.stringify({
question:question
})

});

const data = await response.json();

document.getElementById("response").innerHTML =
data.reply;

}