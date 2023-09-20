document.addEventListener("DOMContentLoaded", function () {
    const studentSelect = document.getElementById('studentSelect');
    const attendanceTypeSelect = document.getElementById('attendanceType');
    const submitAttendanceButton = document.getElementById('submitAttendance');
    const attendanceLog = document.getElementById('attendanceLog');
    const downloadButton = document.getElementById('downloadButton'); // Botón para descargar en Excel
    let attendanceData = [];

    // Recuperar los datos de asistencia almacenados en localStorage al cargar la página
    const storedData = localStorage.getItem('attendanceData');
    if (storedData) {
        attendanceData = JSON.parse(storedData);
        // Mostrar las asistencias almacenadas en la página
        displayAttendances();
    }

    submitAttendanceButton.addEventListener('click', function () {
        const studentName = studentSelect.value;
        const attendanceType = attendanceTypeSelect.value;
        const dateTime = new Date().toLocaleString();

        if (studentName) {
            const entry = document.createElement("p");
            entry.textContent = `${studentName} - ${attendanceType} - ${dateTime}`;
            attendanceLog.appendChild(entry);

            attendanceData.push({ StudentName: studentName, AttendanceType: attendanceType, DateTime: dateTime });

            studentSelect.value = "";
            attendanceTypeSelect.value = "entrada";

            // Almacenar los datos de asistencia en localStorage
            localStorage.setItem('attendanceData', JSON.stringify(attendanceData));
        } else {
            alert("Por favor, selecciona un estudiante.");
        }
    });

    // Función para mostrar las asistencias almacenadas en la página
    function displayAttendances() {
        attendanceLog.innerHTML = "";
        for (const entry of attendanceData) {
            const p = document.createElement("p");
            p.textContent = `${entry.StudentName} - ${entry.AttendanceType} - ${entry.DateTime}`;
            attendanceLog.appendChild(p);
        }
    }

    // Función para descargar los datos de asistencia en un archivo de Excel
    const downloadAttendanceAsExcel = () => {
        if (attendanceData.length === 0) {
            alert("No hay asistencias para descargar.");
            return;
        }

        const XLSX = require('xlsx');
        const worksheet = XLSX.utils.json_to_sheet(attendanceData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Asistencias');

        // Crear un blob con el archivo Excel
        const blob = XLSX.write(workbook, { bookType: 'xlsx', type: 'blob' });

        // Crear un objeto URL para el blob
        const url = URL.createObjectURL(blob);

        // Crear un enlace de descarga y simular un clic
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'asistencias.xlsx';
        document.body.appendChild(a);
        a.click();

        // Liberar el objeto URL
        URL.revokeObjectURL(url);
    };

    // Agregar un evento al botón de descarga para llamar a la función de Excel
    downloadButton.addEventListener('click', downloadAttendanceAsExcel);

    // Agregar un evento para evitar la recarga accidental de la página
    window.addEventListener('beforeunload', function (event) {
        if (attendanceData.length > 0) {
            event.preventDefault();
            event.returnValue = 'Tienes asistencias no guardadas. ¿Seguro que quieres salir?';
        }
    });
});
