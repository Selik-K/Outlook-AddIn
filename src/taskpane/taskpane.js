/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    //document.addEventListener(".share").onclick =
  }
});

/*
 *urlfy method is used to extract urls from Text/string
 */
function urlify(text) {
  var urlRegex = /(https?:\/\/[^\s]+)/g;
  return text.match(urlRegex);
}

/*
 *  list of links for each social medias currently only 3 of em
 */
var twitterLinks = [];
var linkedInLinks = [];
var facebookLinks = [];

/*
 *  run method is the main method which will run by clicking run button
 */
export async function run() {
  Office.context.mailbox.item.body.getAsync("text", function callback(result) {
    let HtmlBody = "";

    try {
      let listOfUrls = urlify(result.value);
      if (listOfUrls != null && listOfUrls.length > 0) {
        listOfUrls.forEach(extractLinks);
        HtmlBody = CreateListOfLinks();
      } else {
        HtmlBody = "Couldn't find any links for SocialMedia.";
      }

      document.getElementById("item-subject").innerHTML = HtmlBody;

      (function (d, s, id) {
        var js,
          fjs = d.getElementsByTagName(s)[0];
        if (d.getElementById(id)) return;
        js = d.createElement(s);
        js.id = id;
        js.src = "https://connect.facebook.net/en_US/sdk.js#xfbml=1&version=v3.0";
        fjs.parentNode.insertBefore(js, fjs);
      })(document, "script", "facebook-jssdk");
    } catch (error) {
      document.getElementById("item-subject").innerHTML = error.message;
    }
  });
}

//adding links to different arrays
function extractLinks(item) {
  if (item.includes("twitter.com")) twitterLinks.push(item);
  if (item.includes("linkedin.com")) linkedInLinks.push(item);
  if (item.includes("facebook.com")) facebookLinks.push(item);
}

function CreateListOfLinks() {
  let text = "";

  if (linkedInLinks.length > 0) {
    text += "<ul>";
    text += "<h3>LinkedIn</h3>";
    linkedInLinks.forEach((item) => {
      text +=
        "<li>" +
        item +
        "<strong> linkedIn</strong>" +
        `<a  href="https://www.linkedin.com/sharing/share-offsite/?url=${item}"  target="popup" 
        onclick="window.open('https://www.linkedin.com/sharing/share-offsite/?url=${item}','popup','width=600,height=600'); return false;">Share</a>
      </li>`;
    });
    text += "</ul>";
  }

  if (facebookLinks.length > 0) {
    text += "<ul>";
    text += "<h3>facebook</h3>";
    facebookLinks.forEach((item) => {
      text +=
        "<li>" +
        item +
        "<strong> Facebook</strong>" +
        `<div class="fb-share-button" 
            data-href="${item}" 
            data-layout="button_count">
            </div>
            </li>`;
    });

    if (twitterLinks.length > 0) {
      text += "<ul>";
      text += "<h3>Twitter</h3>";
      twitterLinks.forEach((item) => {
        text +=
          "<li>" +
          item +
          " <strong> Twitter</strong>" +
          `<a  href="https://twitter.com/share?text=Share&url=${item}">Share</a>  
        </li>`;
      });
      text += "</ul>";
    }
    text += "</ul>";
  }

  if (linkedInLinks.length == 0 && twitterLinks.length == 0 && facebookLinks.length == 0) {
    text = "<h2>Couldn't find any links for SocialMedia.</h2>";
  }
  twitterLinks = [];
  linkedInLinks = [];
  facebookLinks = [];

  return text;
}
