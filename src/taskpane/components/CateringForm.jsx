import React, { useState, useEffect } from "react";
import { Button, makeStyles } from "@fluentui/react-components";
import { ArrowClockwise24Regular } from "@fluentui/react-icons";
import { MessageBar } from "@fluentui/react-components";
import form from "form-urlencoded";

// Pas deze aan naar jouw eigen site & lijst (let op de trailing colon bij de sitePath)
const sitePath = "20200213bvlofc201315.sharepoint.com:/sites/H101-FAC-Voedinsgdienst:";
const listId = "57642914-fce0-4ab7-8d47-1434d8964cc7";

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
    padding: "12px 24px",
    border: "none",
    borderRadius: "6px",
    cursor: "pointer",
    fontSize: "16px",
    transition: "background-color 0.3s, transform 0.2s",
    "&:hover": {
      backgroundColor: "#004d4d",
      transform: "scale(1.02)",
    },
    "&:active": {
      transform: "scale(0.98)",
    },
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

  // Outlook appointment data: sla zowel de ruwe als de weergavewaarde op
  const [appointmentData, setAppointmentData] = useState({
    subject: "Laden...",
    location: "Laden...",
    start: "Laden...", // Weergavewaarde (locale string)
    startRaw: null,    // Ruwe datumwaarde
    end: "Laden...",
  });
  const [loading, setLoading] = useState(true);

  // Catering data
  const [cateringData, setCateringData] = useState({
    aantalPersonen: "",
    opmerkingen: "",
    opstelling: "Standaard",
    andereOpstelling: "",
  });

  // Voor notificaties (bijv. MessageBar van FluentUI)
  const [notification, setNotification] = useState(null);

  useEffect(() => {
    fetchAppointmentData();
  }, []);

  const fetchAppointmentData = () => {
    setLoading(true);
    try {
      const item = Office.context.mailbox.item;
      if (!item) {
        console.warn("Geen item-object gevonden in mailbox.");
        setAppointmentData({
          subject: "Geen afspraak geselecteerd",
          location: "Geen afspraak geselecteerd",
          start: "Geen afspraak geselecteerd",
          startRaw: null,
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

      // Haal starttijd op: sla zowel de ruwe waarde als de weergavewaarde op
      if (item.start && item.start.getAsync) {
        item.start.getAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const rawValue = result.value;
            const displayValue = new Date(rawValue).toLocaleString();
            setAppointmentData((prev) => ({
              ...prev,
              start: displayValue,
              startRaw: rawValue,
            }));
          } else {
            setAppointmentData((prev) => ({ ...prev, start: "Onbekend", startRaw: null }));
          }
        });
      } else {
        const rawValue = item.start ? item.start : null;
        const displayValue = rawValue ? new Date(rawValue).toLocaleString() : "Onbekend";
        setAppointmentData((prev) => ({ ...prev, start: displayValue, startRaw: rawValue }));
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
        startRaw: null,
        end: "Fout bij ophalen",
      });
    } finally {
      setLoading(false);
    }
  };

  const handleCateringChange = (field, value) => {
    setCateringData((prev) => ({ ...prev, [field]: value }));
  };

  // Gebruik de submit-knop ("Aanvraag Versturen") om alle gegevens via Graph te verzenden
  const handleSubmit = async (e) => {
    e.preventDefault();
    console.log("Afspraakgegevens:", appointmentData);
    console.log("Cateringgegevens:", cateringData);
    try {
      const graphToken = await getGraphTokenWithFallback();
      console.log("Ontvangen Graph token:", graphToken);

      // Converteer de ruwe starttijd naar een ISO-string, indien geldig
      const isoStart =
        appointmentData.startRaw && !isNaN(new Date(appointmentData.startRaw).getTime())
          ? new Date(appointmentData.startRaw).toISOString()
          : "";

      const itemFields = {
        Title: appointmentData.subject || "Geen onderwerp",
        Zaal: appointmentData.location || "Onbekend",
        Datum: isoStart,
        Aantal: cateringData.aantalPersonen ? parseInt(cateringData.aantalPersonen, 10) : 0,
        Opmerking: cateringData.opmerkingen || "",
        Opstelling: cateringData.opstelling || "",
      };

      if (cateringData.opstelling === "Andere") {
        itemFields.AndereOpstelling = cateringData.andereOpstelling || "";
      }

      const endpoint = `https://graph.microsoft.com/v1.0/sites/${sitePath}/lists/${listId}/items`;
      const requestBody = { fields: itemFields };

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${graphToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error("Fout bij toevoegen item: " + errorText);
      }

      setNotification({
        message: "Aanvraag werd doorgestuurd!",
        type: "success",
      });
    } catch (error) {
      console.error("Fout bij het versturen van het item:", error);
      alert("Er is een fout opgetreden: " + error.message);
    }
  };

  // Fallback-authenticatie: opent een Office-dialog waarin MSAL de token ophaalt
  const fallbackAuth = () => {
    return new Promise((resolve, reject) => {
      const dialogUrl = window.location.origin + "/fallbackauthdialog.html";
      Office.context.ui.displayDialogAsync(
        dialogUrl,
        { height: 60, width: 30 },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            return reject(asyncResult.error);
          }
          const dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            const message = JSON.parse(args.message);
            if (message.status === "success") {
              dialog.close();
              resolve(message.result);
            } else {
              dialog.close();
              reject(message.error || message.result);
            }
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            reject("Dialog gesloten door de gebruiker.");
          });
        }
      );
    });
  };

  // Probeer eerst SSO, anders gebruik fallback
  const getGraphTokenWithFallback = async () => {
    try {
      const bootstrapToken = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true,
      });
      console.log("Office SSO-token opgehaald, maar geen server OBO => fallback");
      return fallbackAuth();
    } catch (err) {
      console.warn("SSO mislukt of niet ondersteund, fallback auth wordt gebruikt:", err);
      return fallbackAuth();
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
            {loading
              ? "Laden..."
              : appointmentData.start ||
                (appointmentData.startRaw &&
                  new Date(appointmentData.startRaw).toLocaleString())}
          </span>
        </div>
        <div className={styles.appointmentRowLast}>
          <span className={styles.appointmentLabel}>Eindtijd:</span>
          <span className={styles.appointmentValue}>
            {loading ? "Laden..." : appointmentData.end}
          </span>
        </div>
      </div>

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
            Aanvraag Versturen
          </button>
        </div>
      </form>

      {notification && (
        <div style={{ marginTop: "20px" }}>
          <MessageBar
            appearance={notification.type}
            onDismiss={() => setNotification(null)}
            dismissButtonAriaLabel="Sluiten"
          >
            {notification.message}
          </MessageBar>
        </div>
      )}
    </div>
  );
};

export default CateringForm;
