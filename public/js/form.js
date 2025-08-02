document.addEventListener('DOMContentLoaded', () => {
    // Fecha actual
    const fechaInput = document.getElementById('fechaEntrega');
    fechaInput.value = new Date().toLocaleDateString('es-CL');

    let selectedColaborador = {};
    const input = document.getElementById('colaboradorInput');
    const list = document.getElementById('suggestions');
    const form = document.getElementById('actaForm');

    input.addEventListener('input', async () => {
        const q = input.value;
        if (!q) {
            list.innerHTML = '';
            return;
        }
        const res = await fetch(`/api/colaboradores?search=${encodeURIComponent(q)}`);
        const arr = await res.json();
        list.innerHTML = arr.map(c => `
      <button type="button" class="list-group-item list-group-item-action"
              data-id_usuario="${c.id_usuario}" data-area="${c.area}"
              data-cargo="${c.cargo}" data-correo="${c.correo}">
        ${c.nombre}
      </button>
    `).join('');
    });

    list.addEventListener('click', e => {
        if (e.target.matches('button')) {
            selectedColaborador = {
                id_usuario: e.target.dataset.id_usuario,
                nombre: e.target.textContent,
                area: e.target.dataset.area,
                cargo: e.target.dataset.cargo,
                correo: e.target.dataset.correo
            };
            input.value = selectedColaborador.nombre;
            document.getElementById('area').value = selectedColaborador.area;
            document.getElementById('cargo').value = selectedColaborador.cargo;
            document.getElementById('correo').value = selectedColaborador.correo;
            list.innerHTML = '';
        }
    });

    form.addEventListener('submit', async e => {
        e.preventDefault();
        if (!selectedColaborador.id_usuario) {
            return alert('Seleccione un colaborador válido');
        }

        const data = {
            nombreEquipo: document.getElementById('nombreEquipo').value,
            tipoEquipo: document.getElementById('tipoEquipo').value,
            colaborador: selectedColaborador,
            usoEquipo: document.getElementById('usoEquipo').value,
            marca: document.getElementById('marca').value,
            modelo: document.getElementById('modelo').value,
            serie: document.getElementById('serie').value,
            hostname: document.getElementById('hostname').value,
            imei: document.getElementById('imei').value,
            otrosDetalle: document.getElementById('otrosDetalle').value,
            estadoEquipo: document.getElementById('estadoEquipo').value,
            comentarios: document.getElementById('comentarios').value,
            elementosAdicionales: document.getElementById('elementosAdicionales').value
        };

        const res = await fetch('/generate-excel', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });

        if (!res.ok) {
            return alert('Hubo un error al generar el archivo');
        }

        // Abrir el Excel
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        window.open(url);

        // ——— Aquí limpiamos el formulario ———
        form.reset();                         // limpia todos los campos
        selectedColaborador = {};             // resetea la selección
        list.innerHTML = '';                  // limpia sugerencias

        // como form.reset quita el valor de fecha, lo volvemos a fijar:
        fechaInput.value = new Date().toLocaleDateString('es-CL');
    });
});
