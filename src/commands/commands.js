let config;
let btnEvent;

// Inicializačná funkcia prebehne pri každom spustení stránky.
Office.initialize = function () {};

function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "github-error",
    {
      type: "errorMessage",
      message: error,
    },
    function (result) {}
  );
}

let settingsDialog;

function insertDefaultGist(event) {
  config = getConfig();

  // Skontroluje, či bol doplnok nakonfigurovaný.
  if (config && config.defaultGistId) {
    // Získa predvolený obsah gistu a vložího ho do správy.
    try {
      getGist(config.defaultGistId, function (gist, error) {
        if (gist) {
          buildBodyContent(gist, function (content, error) {
            if (content) {
              Office.context.mailbox.item.body.setSelectedDataAsync(
                content,
                { coercionType: Office.CoercionType.Html },
                function (result) {
                  event.completed();
                }
              );
            } else {
              showError(error);
              event.completed();
            }
          });
        } else {
          showError(error);
          event.completed();
        }
      });
    } catch (err) {
      showError(err);
      event.completed();
    }
  } else {
    // Uloží objekt udalosti, aby sme ho mohli dokončiť neskôr.
    btnEvent = event;
    // Ak ešte doplnok nie je nakonfigurovaný, zobrazí sa dialógové okno nastaveniami
    // warn=1 zobrazí sa varovanie – potrebná konfigurácia doplnku pomocou npm príkazu v príkazovom riadku.

    const url = new URI("dialog.html?warn=1").absoluteTo(window.location).toString();
    const dialogOptions = { width: 20, height: 40, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function (result) {
      settingsDialog = result.value;
      settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
      settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
    });
  }
}

// Zaregistrujeme funkciu.
Office.actions.associate("insertDefaultGist", insertDefaultGist);

function receiveMessage(message) {
  config = JSON.parse(message.message);
  setConfig(config, function (result) {
    settingsDialog.close();
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
  });
}

function dialogClosed(message) {
  settingsDialog = null;
  btnEvent.completed();
  btnEvent = null;
}
