(function () {
  "use strict";

  // Funkcia inicializácia Office sa musí spustiť pri každom načítaní novej stránky.
  Office.initialize = function (reason) {
    jQuery(document).ready(function () {
      if (window.location.search) {
        // Skontroluje, či sa má zobraziť upozornenie.
        const warn = getParameterByName("warn");
        if (warn) {
          $(".not-configured-warning").show();
        } else {
          // Zistí, či boli odovzdané konfiguračné hodnoty.
          // Ak áno, vopred vyplní hodnoty.
          const user = getParameterByName("gitHubUserName");
          const gistId = getParameterByName("defaultGistId");

          $("#github-user").val(user);
          loadGists(user, function (success) {
            if (success) {
              $(".ms-ListItem").removeClass("is-selected");
              $("input")
                .filter(function () {
                  return this.value === gistId;
                })
                .addClass("is-selected")
                .attr("checked", "checked");
              $("#settings-done").removeAttr("disabled");
            }
          });
        }
      }

      // Pri zmene používateľského mena GitHub,
      // skúsime načítať gists.
      $("#github-user").on("change", function () {
        $("#gist-list").empty();
        const ghUser = $("#github-user").val();
        if (ghUser.length > 0) {
          loadGists(ghUser);
        }
      });

      // Po výbere tlačidla Hotovo odošle
      // hodnoty späť volajúcemu ako serializovaný
      // objekt.
      $("#settings-done").on("click", function () {
        const settings = {};

        settings.gitHubUserName = $("#github-user").val();

        const selectedGist = $(".ms-ListItem.is-selected");
        if (selectedGist) {
          settings.defaultGistId = selectedGist.val();

          sendMessage(JSON.stringify(settings));
        }
      });
    });
  };

  // Načítanie zoznamu gists pre používateľa pomocou rozhrania GitHub API
  // a zostaví zoznam.
  function loadGists(user, callback) {
    getUserGists(user, function (gists, error) {
      if (error) {
        $(".gist-list-container").hide();
        $("#error-text").text(JSON.stringify(error, null, 2));
        $(".error-display").show();
        if (callback) callback(false);
      } else {
        $(".error-display").hide();
        buildGistList($("#gist-list"), gists, onGistSelected);
        $(".gist-list-container").show();
        if (callback) callback(true);
      }
    });
  }

  function onGistSelected() {
    $(".ms-ListItem").removeClass("is-selected").removeAttr("checked");
    $(this).children(".ms-ListItem").addClass("is-selected").attr("checked", "checked");
    $(".not-configured-warning").hide();
    $("#settings-done").removeAttr("disabled");
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    const regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return "";
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
