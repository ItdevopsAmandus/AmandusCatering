import React, { useState, useEffect } from "react";
import { Button, makeStyles } from "@fluentui/react-components";
import { ArrowClockwise24Regular } from "@fluentui/react-icons";
import form from "form-urlencoded";

// Vervang deze waarden door jouw eigen gegevens
const CLIENT_ID = "82d99688-d922-4bfc-8d2d-e2871eb05ebd";
const CLIENT_SECRET = "YHW8Q~-iictXaRYBw~PV5U_9lF_.bCS1KLUMsc.W";
const TENANT_ID = "82022306-deb0-41be-94c4-763bf46d3547";

// Fallback-authenticatie: open een Office-dialog voor MSAL-login
const fallbackAuth = async () => {
  return new Promise((resolve, reject) => {
    const url = "/fallbackauthdialog.html";
    const fullUrl =
      location.protocol +
      "//" +
      location.hostname +
      (location.port ? ":" + location.port : "") +
      url;
    Office.context.ui.displayDialogAsync(
      fullUrl,
      { height: 60, width: 30 },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(asyncResult.error);
          return;
        }
        const loginDialog = asyncResult.value;
        loginDialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg) => {
            const messageFromDialog = JSON.parse(arg.message);
            if (messageFromDialog.status === "success") {
              loginDialog.close();
              resolve(messageFromDialog.result);
            } else {
              loginDialog.close();
              reject(messageFromDialog.error || messageFromDialog.result);
            }
          }
        );
      }
    );
  });
};

// Helperfunctie die eerst SSO probeert en zo niet de fallback gebruikt.
// Vervolgens ruilt deze de verkregen token in voor een Graph-token via de on‑behalf‑of flow.
const getGraphTokenWithFallback = async () => {
  let ssoToken;
  try {
    ssoToken = await Office.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: true,
    });
  } catch (error) {
    console.warn(
      "Office.auth.getAccessToken wordt niet ondersteund, fallback auth gebruiken",
      error
    );
    ssoToken = await fallbackAuth();
  }
  // Verkrijg het 'assertion'-gedeelte uit het Bearer token
  const [, assertion] = `Bearer ${ssoToken}`.split(" ");
  const formParams = {
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    assertion: assertion,
    requested_token_use: "on_behalf_of",
    scope: "https://graph.microsoft.com/.default",
  };
  const encodedForm = form(formParams);
  const tokenURL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const tokenResponse = await fetch(tokenURL, {
    method: "POST",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: encodedForm,
  });
  const tokenJson = await tokenResponse.json();
  if (tokenJson.error) {
    throw new Error(`Token request failed: ${tokenJson.error_description}`);
  }
  return tokenJson.access_token;
};

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "20px",
    maxWidth: "600px",
    margin: "20px auto",
    backgroundColor: "#faf9f8",
    borderRadius: "8px",
    boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
  },
  appointmentCard: {
    width: "100%",
    backgroundColor: "#fff",
    padding: "15px",
    borderRadius: "8px",
    boxShadow: "0 1px 4px rgba(0,0,0,0.1)",
    marginBottom: "20px",
  },
  appointmentRow: {
    display: "flex",
    justifyContent: "space-between",
    padding: "5px 0",
    borderBottom: "1px solid #eee",
    fontSize: "14px",
  },
  appointmentRowLast: {
    display: "flex",
    justifyContent: "space-between",
    padding: "5px 0",
    fontSize: "14px",
  },
  appointmentLabel: {
    fontWeight: "600",
    color: "#555",
    marginRight: "10px",
  },
  appointmentValue: {
    color: "#333",
    textAlign: "right",
  },
  refreshButton: {
    marginBottom: "15px",
    backgroundColor: "#008075",
    color: "white",
    "&:hover": { backgroundColor: "#006a5a" },
  },
  section: {
    width: "100%",
    marginBottom: "20px",
  },
  sectionHeader: {
    marginBottom: "8px",
    fontSize: "18px",
    fontWeight: "600",
    color: "#008075",
  },
  fieldLabel: {
    marginBottom: "4px",
    fontSize: "14px",
    color: "#333",
  },
  inputField: {
    backgroundColor: "white",
    border: "1px solid #b2dfdb",
    padding: "8px",
    borderRadius: "6px",
    width: "100%",
    fontSize: "14px",
    color: "#333",
  },
  textArea: {
    backgroundColor: "white",
    border: "1px solid #b2dfdb",
    padding: "8px",
    borderRadius: "6px",
    width: "100%",
    fontSize: "14px",
    color: "#333",
    minHeight: "60px",
    resize: "vertical",
  },
  submitButton: {
    backgroundColor: "#006a5a",
    color: "white",
    padding: "10px 20px",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "16px",
  },
  testButton: {
    marginTop: "20px",
    backgroundColor: "#0078d4",
    color: "white",
    padding: "10px 20px",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "16px",
  },
});

const CateringForm = () => {
  const styles = useStyles();

  // Afspraakgegevens (uit Outlook)
  const [appointmentData, setAppointmentData] = useState({
    subject: "Laden...",
    location: "Laden...",
    start: "Laden...",
    end: "Laden...",
  });
  const [loading, setLoading] = useState(true);

  // Cateringgegevens die later verstuurd moeten worden
  const [cateringData, setCateringData] = useState({
    aantalPersonen: "",
    opmerkingen: "",
    opstelling: "Standaard",
    andereOpstelling: "",
  });

  useEffect(() => {
    fetchAppointmentData();
  }, []);

  const fetchAppointmentData = () => {
    setLoading(true);
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        console.warn("Geen item-object gevonden.");
        setAppointmentData({
          subject: "Geen afspraak geselecteerd",
          location: "Geen afspraak geselecteerd",
          start: "Geen afspraak geselecteerd",
          end: "Geen afspraak geselecteerd",
        });
        setLoading(false);
        return;
      }

      // Haal onderwerp op
      if (item.subject && item.subject.getAsync) {
        item.subject.getAsync((result) => {
          setAppointmentData((prev) => ({
            ...prev,
            subject:
              result.status === Office.AsyncResultStatus.Succeeded
                ? result.value || "Onbekend"
                : "Onbekend",
          }));
        });
      } else {
        const subject = item.subject || "Onbekend";
        setAppointmentData((prev) => ({ ...prev, subject }));
      }

      // Haal locatie op
      if (item.location && item.location.getAsync) {
        item.location.getAsync((result) => {
          setAppointmentData((prev) => ({
            ...prev,
            location:
              result.status === Office.AsyncResultStatus.Succeeded
                ? result.value || "Onbekend"
                : "Onbekend",
          }));
        });
      } else {
        let location = "Onbekend";
        if (item.location) {
          if (typeof item.location === "string") {
            location = item.location || "Onbekend";
          } else if (typeof item.location === "object" && item.location.displayName) {
            location = item.location.displayName || "Onbekend";
          }
        }
        setAppointmentData((prev) => ({ ...prev, location }));
      }

      // Haal starttijd op
      if (item.start && item.start.getAsync) {
        item.start.getAsync((result) => {
          setAppointmentData((prev) => ({
            ...prev,
            start:
              result.status === Office.AsyncResultStatus.Succeeded
                ? new Date(result.value).toLocaleString()
                : "Onbekend",
          }));
        });
      } else {
        const start = item.start ? new Date(item.start).toLocaleString() : "Onbekend";
        setAppointmentData((prev) => ({ ...prev, start }));
      }

      // Haal eindtijd op
      if (item.end && item.end.getAsync) {
        item.end.getAsync((result) => {
          setAppointmentData((prev) => ({
            ...prev,
            end:
              result.status === Office.AsyncResultStatus.Succeeded
                ? new Date(result.value).toLocaleString()
                : "Onbekend",
          }));
        });
      } else {
        const end = item.end ? new Date(item.end).toLocaleString() : "Onbekend";
        setAppointmentData((prev) => ({ ...prev, end }));
      }
    } catch (error) {
      console.error("Fout bij ophalen van afspraakgegevens:", error);
      setAppointmentData({
        subject: "Fout bij ophalen",
        location: "Fout bij ophalen",
        start: "Fout bij ophalen",
        end: "Fout bij ophalen",
      });
    } finally {
      setLoading(false);
    }
  };

  const handleCateringChange = (field, value) => {
    setCateringData((prev) => ({ ...prev, [field]: value }));
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    console.log("Afspraakgegevens:", appointmentData);
    console.log("Cateringgegevens:", cateringData);
    alert("Gegevens verstuurd (check console)");
  };

  // Functie om een dummy item naar de SharePoint-lijst te sturen
  const sendDummyItemToList = async () => {
    try {
      // Haal Graph-token op via de helper met fallback
      const graphToken = await getGraphTokenWithFallback();

      // Endpoint: gebruik pad-based URL voor jouw site en list-id
      const sitePath = "20200213bvlofc201315.sharepoint.com:/sites/H101-FAC-Voedinsgdienst";
      const listId = "57642914-fce0-4ab7-8d47-1434d8964cc7";
      const endpoint = `https://graph.microsoft.com/v1.0/sites/${sitePath}/lists/${listId}/items`;

      // Dummy data voor het item (pas de veldnamen aan indien nodig)
      const dummyItem = {
        fields: {
          Title: "Dummy item via SSO-test",
        },
      };

      // Verstuur de POST-aanvraag naar de Graph API
      const listResponse = await fetch(endpoint, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${graphToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(dummyItem),
      });

      if (listResponse.ok) {
        alert("Dummy item succesvol toegevoegd aan de lijst!");
      } else {
        const errorText = await listResponse.text();
        alert("Fout bij toevoegen item: " + errorText);
      }
    } catch (error) {
      console.error("Fout bij het versturen van het dummy item:", error);
      alert("Er is een fout opgetreden: " + error.message);
    }
  };

  return (
    <div className={styles.container}>
      <Button
        className={styles.refreshButton}
        icon={<ArrowClockwise24Regular />}
        onClick={fetchAppointmentData}
      >
        Ververs Afspraakgegevens
      </Button>

      {/* Afspraakgegevens kaart */}
      <div className={styles.appointmentCard}>
        <div className={styles.appointmentRow}>
          <span className={styles.appointmentLabel}>Onderwerp:</span>
          <span className={styles.appointmentValue}>
            {loading ? "Laden..." : appointmentData.subject}
          </span>
        </div>
        <div className={styles.appointmentRow}>
          <span className={styles.appointmentLabel}>Locatie:</span>
          <span className={styles.appointmentValue}>
            {loading ? "Laden..." : appointmentData.location}
          </span>
        </div>
        <div className={styles.appointmentRow}>
          <span className={styles.appointmentLabel}>Starttijd:</span>
          <span className={styles.appointmentValue}>
            {loading ? "Laden..." : appointmentData.start}
          </span>
        </div>
        <div className={styles.appointmentRowLast}>
          <span className={styles.appointmentLabel}>Eindtijd:</span>
          <span className={styles.appointmentValue}>
            {loading ? "Laden..." : appointmentData.end}
          </span>
        </div>
      </div>

      {/* Cateringformulier */}
      <form onSubmit={handleSubmit} className={styles.section}>
        <div className={styles.sectionHeader}>Catering Gegevens</div>
        <div className={styles.fieldLabel}>Aantal Personen (Koffie & Thee)</div>
        <input
          type="number"
          min="0"
          value={cateringData.aantalPersonen}
          onChange={(e) => handleCateringChange("aantalPersonen", e.target.value)}
          className={styles.inputField}
          placeholder="Bijv. 10"
        />
        <div className={styles.fieldLabel} style={{ marginTop: "10px" }}>
          Opmerkingen
        </div>
        <textarea
          value={cateringData.opmerkingen}
          onChange={(e) => handleCateringChange("opmerkingen", e.target.value)}
          className={styles.textArea}
          placeholder="Eventuele extra wensen of opmerkingen"
        />
        <div className={styles.fieldLabel} style={{ marginTop: "10px" }}>
          Opstelling
        </div>
        <select
          value={cateringData.opstelling}
          onChange={(e) => handleCateringChange("opstelling", e.target.value)}
          className={styles.inputField}
        >
          <option value="Standaard">Standaard</option>
          <option value="U-vorm">U-vorm</option>
          <option value="Cirkel">Cirkel</option>
          <option value="Klas">Klas</option>
          <option value="Andere">Andere</option>
        </select>
        {cateringData.opstelling === "Andere" && (
          <>
            <div className={styles.fieldLabel} style={{ marginTop: "10px" }}>
              Specificatie Opstelling
            </div>
            <input
              type="text"
              value={cateringData.andereOpstelling}
              onChange={(e) => handleCateringChange("andereOpstelling", e.target.value)}
              className={styles.inputField}
              placeholder="Specificeer de gewenste opstelling"
            />
          </>
        )}
        <div style={{ textAlign: "center", marginTop: "20px" }}>
          <button type="submit" className={styles.submitButton}>
            Gegevens Versturen
          </button>
        </div>
      </form>

      {/* Knop om een dummy item naar de SharePoint-lijst te sturen */}
      <button className={styles.testButton} onClick={sendDummyItemToList}>
        Test Dummy Item Versturen
      </button>
    </div>
  );
};

export default CateringForm;
