## Patient Registration System

This is a Python application for registering patient information using a graphical user interface (GUI). It allows users to input patient details, select the type of demand, add time intervals for conviviality, and save the registration information to an Excel file.

### Dependencies

The following libraries are required to run the application:

- `sys`
- `os`
- `openpyxl`
- `PyQt5`
- `datetime`

You can install the dependencies using pip:

```
pip install openpyxl PyQt5
```

### Usage

To run the application, execute the following command:

```
python patient_registration.py
```

The Patient Registration window will appear, where you can input patient details and select the type of demand. After filling in the necessary information, click the "Registrar" button to save the registration.

#### Patient Details

- **Nome do Paciente**: Enter the name of the patient in this field.

#### Type of Demand

Select one or more types of demand for the patient. The available options are:

- A
- R
- M
- AN
- C
- RM
- Grupos/Eventos
- Outros
- AI
- REA

The checkboxes represent different demand types. You can select multiple checkboxes to indicate multiple types of demand.

#### Time Intervals

If you select the "C" (Convivência) checkbox, you can specify the time interval for conviviality. Click on the checkbox to enable the time input fields for the start and end time of the interval. Enter the desired time in the format "HH:mm" (e.g., 09:00) and press "Enter" or click outside the input field to apply the changes.

#### Tipo de Encaminhamento

If you select either "AI" or "REA" checkbox, a dialog box will appear to choose the type of encaminhamento. Select the desired option from the drop-down menu and click "Salvar" to save the selection.

#### Prof. de referência

Enter the reference professional or attending professional in this field.

#### Almoço, Lanche, Janta

You can select the checkboxes for "Almoço" (Lunch), "Lanche" (Snack), and "Janta" (Dinner) if applicable. These checkboxes indicate whether the patient requires these meals.

#### Observações

You can enter any additional observations or notes about the patient in this field.

### Data Storage

The application stores patient registration data in an Excel file named "pacientes_recepção.xlsx". If the file does not exist, it will be created automatically. The file contains the following sheets:

- **Pacientes**: Contains registration information for patients.
- **Almoço**: Contains registration information for patients requiring lunch.
- **Lanche**: Contains registration information for patients requiring a snack.
- **Janta**: Contains registration information for patients requiring dinner.
- **Acolhimentos**: Contains registration information for patients with demand types "AI" or "REA".
- **Consolidados**: Contains consolidated data for each day, including the number of patients for lunch, snack, dinner, and total patients.
- **Consolidados Totais**: Contains the total count of patients for lunch, snack, dinner, and overall.

The sheets are automatically filtered based on the current date, and the data is updated accordingly.







### Note

This code assumes that an image file named "OIG.jpeg" exists in the same directory as the Python script and is used for displaying the logo.
