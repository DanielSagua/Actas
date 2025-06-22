document.addEventListener('DOMContentLoaded', () => {
    const fechaInput = document.getElementById('fechaEntrega');
    fechaInput.value = new Date().toLocaleDateString('es-CL');

    let selectedColaborador = {};
    const input = document.getElementById('colaboradorInput');
    const list = document.getElementById('suggestions');
    const form = document.getElementById('devolucionForm');

    input.addEventListener('input', async () => {
        const q = input.value;
        if (!q) { list.innerHTML = ''; return; }
        const res = await fetch(`/api/colaboradores?search=${encodeURIComponent(q)}`);
        const arr = await res.json();
        list.innerHTML = arr.map(c => `
      <button type="button" class="list-group-item list-group-item-action"
              data-id="${c.id}" data-area="${c.area}"
              data-cargo="${c.cargo}" data-correo="${c.correo}">
        ${c.nombre}
      </button>
    `).join('');
    });

    list.addEventListener('click', e => {
        if (e.target.matches('button')) {
            selectedColaborador = {
                id: e.target.dataset.id,
                nombre: e.target.textContent.trim(),
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
        if (!selectedColaborador.id) return alert('Seleccione un colaborador válido');
        const data = {
            nombreEquipo: document.getElementById('nombreEquipo').value.trim(),
            tipoEquipo: document.getElementById('tipoEquipo').value,
            colaborador: selectedColaborador,
            marca: document.getElementById('marca').value.trim(),
            modelo: document.getElementById('modelo').value.trim(),
            serie: document.getElementById('serie').value.trim(),
            hostname: document.getElementById('hostname').value.trim(),
            imei: document.getElementById('imei').value.trim(),
            otrosDetalle: document.getElementById('otrosDetalle').value.trim(),
            estadoEquipo: document.getElementById('estadoEquipo').value.trim(),
            comentarios: document.getElementById('comentarios').value.trim(),
            elementosAdicionales: document.getElementById('elementosAdicionales').value.trim()
        };
        const res = await fetch('/generate-excel-devolucion', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        if (!res.ok) return alert('Error al generar el archivo de devolución');
        const blob = await res.blob();
        window.open(URL.createObjectURL(blob));
        form.reset();
        selectedColaborador = {};
        list.innerHTML = '';
        fechaInput.value = new Date().toLocaleDateString('es-CL');
    });
});