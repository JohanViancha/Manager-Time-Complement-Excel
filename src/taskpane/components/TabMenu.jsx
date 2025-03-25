import {
  Divider,
  Button,
  Field,
  makeStyles,
  Select,
  Tab,
  Tag,
  TabList,
} from "@fluentui/react-components";
import { DatePicker, defaultDatePickerStrings } from "@fluentui/react-datepicker-compat";
import {
  bundleIcon,
  TextBulletListSquareClockFilled,
  TextBulletListSquareClockRegular,
  WindowBulletListAddFilled,
  WindowBulletListAddRegular,
} from "@fluentui/react-icons";
import { TimePicker, formatDateToTimeString } from "@fluentui/react-timepicker-compat";
import React from "react";
import { registerActivity, calculateForActivityandCompany } from "../taskpane";

const localizedStrings = {
  ...defaultDatePickerStrings,
  days: ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"],
  shortDays: ["D", "L", "M", "M", "J", "V", "S"],
  months: [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Septiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
  ],

  shortMonths: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
  goToToday: "Ir a hoy",
};

const onFormatDate = (date) => {
  return !date
    ? ""
    : `${localizedStrings.months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
};

const useStyles = makeStyles({
  root: {
    alignItems: "flex-start",
    display: "flex",
    flexDirection: "column",
    justifyContent: "flex-start",
    padding: "50px 20px",
    rowGap: "20px",
  },
});

const useStylesDiv = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    justifyContent: "between",
  },
});

const CalendarMonth = bundleIcon(WindowBulletListAddFilled, WindowBulletListAddRegular);
const Time = bundleIcon(TextBulletListSquareClockFilled, TextBulletListSquareClockRegular);

const TabMenu = () => {
  const styles = useStyles();
  const stylesDiv = useStylesDiv();

  const [selectedValue, setSelectedValue] = React.useState("register");
  const [timeTotal, setTimeTotal] = React.useState(0);
  const [formData, setFormData] = React.useState({
    activity: "Trabajo individual",
    project: "Portón de sol",
    timeStart: formatDateToTimeString(new Date(), { hourCycle: "h24" }),
    time: "30 minutos",
    date: new Date(),
  });

  const [formDataTime, setFormDataTime] = React.useState({
    activity: "Trabajo individual",
    project: "Portón de sol",
    book: "Enero 2025",
  });

  const [defaultSelectedTime] = React.useState(new Date());
  const onTabSelect = (event, data) => {
    setSelectedValue(data.value);
  };

  const handleChange = (e) => {
    setFormData({
      ...formData,
      [e.target.name]: e.target.value,
    });
  };

  const handleChangeTime = (e) => {
    setFormDataTime({
      ...formDataTime,
      [e.target.name]: e.target.value,
    });
  };

  const registerTime = (e) => {
    e.preventDefault();
    registerActivity(formData).then((data) => {});
  };

  const calculateTimeForProject = (e) => {
    e.preventDefault();

    calculateForActivityandCompany(formDataTime).then((data) => {
      setTimeTotal(data);
    });
  };

  const Register = React.memo(() => (
    <form onSubmit={registerTime}>
      <div role="tabpanel" aria-labelledby="Register">
        <Field label="Proyecto" required>
          <Select name="project" value={formData.project} onChange={handleChange}>
            <option value={"CISO Celsa"}>CISO Celsa</option>
            <option value={"Portón de sol"}>Portón de sol</option>
            <option value={"Interno PDP"}>Interno PDP</option>
            <option value={"CISO Rsec"}>CISO Rsec</option>
            <option value={"CISO Caseware"}>CISO Caseware</option>
            <option value={"Entrenamiento"}>Entrenamiento</option>
            <option value={"IAM Lemco"}>IAM Lemco</option>
            <option value={"SGSI PDP"}>SGSI PDP</option>
            <option value={"Personal"}>Personal</option>
            <option value={"Inducción"}>Inducción</option>
          </Select>
        </Field>
        <Field label="Actividad" required>
          <Select name="activity" value={formData.activity} onChange={handleChange}>
            <option value={"Trabajo individual"}>Trabajo individual</option>
            <option value={"Investigación"}>Investigación</option>
            <option value={"Personal"}>Personal</option>
            <option value={"Almuerzo"}>Almuerzo</option>
            <option value={"Permiso"}>Permiso</option>
            <option value={"Vacaciones"}>Vacaciones</option>
            <option value={"Incapacidad"}>Incapacidad</option>
            <option value={"Cierre del día"}>Cierre del día</option>
            <option value={"Reunión Interna"}>Reunión Interna</option>
          </Select>
        </Field>
        <Field label="Fecha" required>
          <DatePicker
            name="date"
            strings={localizedStrings}
            formatDate={onFormatDate}
            className={styles.control}
            onChange={handleChange}
            value={formData.date}
          />
        </Field>
        <Field label="Seleccionar hora de inicio">
          <TimePicker
            startHour={7}
            endHour={19}
            increment={30}
            name="timeStart"
            value={formData.timeStart}
            onChange={handleChange}
          />
        </Field>
        <Field label="Horas trabajadas" required>
          <Select defaultValue={30} name="time" value={formData.time} onChange={handleChange}>
            <option value="30">30 min</option>
            <option value="60">1 hora</option>
            <option value="90">1 hora 30 minutos</option>
            <option value="120">2 horas</option>
            <option value="150">2 horas 30 minutos</option>
            <option value="180">3 horas</option>
            <option value="210">3 horas 30 minutos</option>
            <option value="240">4 horas</option>
          </Select>
        </Field>
        <Button type="submit">Registrar tiempos</Button>
      </div>
    </form>
  ));
  const CalculateTime = React.memo(() => (
    <form onSubmit={calculateTimeForProject} className={stylesDiv.root}>
      <div role="tabpanel" aria-labelledby="calculateTime">
        <Field label={"Seleccionar mes"} required>
          <Select name="book" value={formDataTime.book} onChange={handleChangeTime}>
            {localizedStrings.months.map((month, index) => (
              <option key={index} value={`${month} 2025`}>{`${month} 2025`}</option>
            ))}
          </Select>

          <Field label="Proyecto" required>
            <Select name="project" value={formDataTime.project} onChange={handleChangeTime}>
              <option value={"CISO Celsa"}>CISO Celsa</option>
              <option value={"Portón de sol"}>Portón de sol</option>
              <option value={"Interno PDP"}>Interno PDP</option>
              <option value={"CISO Rsec"}>CISO Rsec</option>
              <option value={"CISO Caseware"}>CISO Caseware</option>
              <option value={"Entrenamiento"}>Entrenamiento</option>
              <option value={"IAM Lemco"}>IAM Lemco</option>
              <option value={"SGSI PDP"}>SGSI PDP</option>
              <option value={"Personal"}>Personal</option>
              <option value={"Inducción"}>Inducción</option>
            </Select>
          </Field>
          <Field label="Actividad" required>
            <Select name="activity" value={formDataTime.activity} onChange={handleChangeTime}>
              <option value={"Trabajo individual"}>Trabajo individual</option>
              <option value={"Investigación"}>Investigación</option>
              <option value={"Personal"}>Personal</option>
              <option value={"Almuerzo"}>Almuerzo</option>
              <option value={"Permiso"}>Permiso</option>
              <option value={"Vacaciones"}>Vacaciones</option>
              <option value={"Incapacidad"}>Incapacidad</option>
              <option value={"Cierre del día"}>Cierre del día</option>
              <option value={"Reunión Interna"}>Reunión Interna</option>
            </Select>
          </Field>
        </Field>
        <Button type="submit">Calcular</Button>
      </div>
    
      <Tag>El total de horas trabajadas es: {timeTotal}</Tag>
    </form>
  ));

  return (
    <div className={styles.root}>
      <TabList selectedValue={selectedValue} onTabSelect={onTabSelect}>
        <Tab id="Register" icon={<CalendarMonth />} value="register">
          Registrar tiempos
        </Tab>
        <Tab id="CalculateTime" icon={<Time />} value="calculateTime">
          Calcular tiempos
        </Tab>
      </TabList>
      <Divider />{" "}
      <div className={styles.panels}>
        {selectedValue === "register" && <Register />}
        {selectedValue === "calculateTime" && <CalculateTime />}
      </div>
    </div>
  );
};

export default TabMenu;
