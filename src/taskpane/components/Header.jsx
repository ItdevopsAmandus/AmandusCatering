import * as React from "react";
import PropTypes from "prop-types";
import { Image, tokens, makeStyles } from "@fluentui/react-components";
import { DrinkCoffee24Regular } from "@fluentui/react-icons";

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
  logo: {
    width: "100px",
    height: "100px",
    marginBottom: "10px",
  },
  titleContainer: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    color: "#008075", // Huisstijl kleur
  },
  title: {
    fontSize: "24px",
    fontWeight: "bold",
    margin: 0,
  },
  icon: {
    fontSize: "28px",
    color: "#008075",
  },
});

const Header = (props) => {
  const { title, logo, message } = props;
  const styles = useStyles();

  return (
    <section className={styles.welcome__header}>
      <Image className={styles.logo} src={logo} alt={title} />
      <div className={styles.titleContainer}>
        <DrinkCoffee24Regular className={styles.icon} />
        <h1 className={styles.title}>{message}</h1>
      </div>
    </section>
  );
};

Header.propTypes = {
  title: PropTypes.string,
  logo: PropTypes.string,
  message: PropTypes.string,
};

export default Header;
