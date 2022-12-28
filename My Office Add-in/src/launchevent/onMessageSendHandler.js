/* global Office */

let _mailboxItem;

function onMessageSendHandler(event) {
  _mailboxItem = Office.context.mailbox.item;

  _mailboxItem.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, function (asyncBody) {
    throwError(asyncBody, "Falha na recuperação do corpo do email: ");

    requisicaoXml(asyncBody, { corpoDoEmail: asyncBody.value });
  });
}
const requisicaoXml = (event, serviceRequest) => {
  const xhr = new XMLHttpRequest();
  xhr.open("POST", `http://localhost:8000/salvaCorpoDoEmail`, true);
  xhr.setRequestHeader("Content-Type", "application/json");

  xhr.onreadystatechange = () => {
    if (xhr.readyState === XMLHttpRequest.DONE && xhr.status === 200) {
      event.asyncContext.completed({ allowEvent: false, errorMessage: "Deu certo." });
    } else {
      _mailboxItem.body.setAsync(JSON.stringify(xhr));
      event.asyncContext.completed({ allowEvent: false, errorMessage: "Deu merda." });
    }
  };

  xhr.send(JSON.stringify(serviceRequest));
};

const throwError = (asyncResult, mensagem = "") => {
  if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
    throw mensagem + JSON.stringify(asyncResult.error);
  }
};

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
