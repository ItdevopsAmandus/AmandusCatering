import React from 'react';

import useSSO from '../../hooks/useSSO';

export default function InsertTextButton() {
  const { getUserInfo } = useSSO();

  const handleInsertText = async () => {
    try {
      const user = await getUserInfo(); 
      // Dit zal eerst SSO proberen. Mislukt dat, dan fallback.

      const textToInsert = `<p>Hallo, ${user.displayName} (${user.mail})!</p>`;
      Office.context.mailbox.item.body.setSelectedDataAsync(
        textToInsert,
        { coercionType: Office.CoercionType.Html },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log('Tekst succesvol ingevoegd');
          } else {
            console.error('Fout bij het invoegen van tekst:', asyncResult.error);
          }
        }
      );
    } catch (error) {
      console.error('Insert Text Error:', error);
    }
  };

  return (
    <button onClick={handleInsertText}>
      Insert Text
    </button>
  );
}
