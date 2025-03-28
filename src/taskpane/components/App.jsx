import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";


import { makeStyles } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular } from "@fluentui/react-icons";


import CateringForm from "./CateringForm";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App = (props) => {
  const { title } = props;
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems = [
    {
      icon: <Ribbon24Regular />,
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: <LockOpen24Regular />,
      primaryText: "Unlock features and functionality",
    },
    {
      icon: <DesignIdeas24Regular />,
      primaryText: "Create and visualize like a pro",
    },
  ];

  return (
    <div className={styles.root}>
     <Header 
  logo="assets/PC_Sint-Amandus_beeldmerk_kleur_rgb.png" 
  title={title} 
  message="Hier kunt u een aanvraag doen voor catering. Controleer of de gegevens correct ingevuld zijn en klik op 'Ververs Afspraakgegevens' indien uw afspraak gewijzigd is." 
/>
     
      <CateringForm/>      
  
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
