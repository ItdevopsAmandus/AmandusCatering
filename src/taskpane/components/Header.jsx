import * as React from "react";
import PropTypes from "prop-types";
import { makeStyles } from "@fluentui/react-components";
import { DrinkCoffee24Regular, Info24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    padding: "40px 20px",
    backgroundColor: "#faf9f8",
    borderRadius: "10px",
    boxShadow: "0 2px 8px rgba(0, 0, 0, 0.1)",
    textAlign: "center",
    width: "100%",
  },
  logoIcon: {
    fontSize: "48px", // Grotere grootte voor het logo
    color: "#008075",
    marginBottom: "10px",
  },
  titleContainer: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: "#008075",
  },
  title: {
    fontSize: "24px",
    fontWeight: "bold",
    margin: 0,
  },
  subtitleBox: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    backgroundColor: "#f3f2f1",
    border: "1px solid #e1dfdd",
    padding: "8px 12px",
    borderRadius: "4px",
    fontSize: "14px",
    color: "#333",
    marginTop: "10px",
    maxWidth: "80%",
    textAlign: "left",
  },
  subtitleIcon: {
    fontSize: "20px",
    color: "#0078d4",
  },
});

const Header = (props) => {
  const { title, message } = props;
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      {/* Gebruik de koffie-icoon als logo */}
      <DrinkCoffee24Regular className={styles.logoIcon} />
      <div className={styles.titleContainer}>
        <h1 className={styles.title}>{title}</h1>
      </div>
      {message && (
        <div className={styles.subtitleBox}>
          <Info24Regular className={styles.subtitleIcon} />
          <span>{message}</span>
        </div>
      )}
    </section>
  );
};

Header.propTypes = {
  title: PropTypes.string,
  message: PropTypes.string,
};

export default Header;
