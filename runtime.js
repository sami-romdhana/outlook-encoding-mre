Office.actions.associate("launchFetch", launchFetch);

function launchFetch(eventObj) {
  fetch("https://api.sampleapis.com/wines/reds")
    .then(function (response) {
      return response.json();
    })
    .then(function (result) {
      return setSignature(result[0].location);
    })
    .catch(function (err) {
      return setSignature(err.toString());
    })
    .finally(function () {
      eventObj.completed();
    });
}

function setSignature(signature) {
  return new Promise(function (resolve) {
    Office.context.mailbox.item.body.setSignatureAsync(
      signature,
      {
        coercionType: "html",
      },
      function () {
        resolve();
      }
    );
  });
}

Office.onReady();
