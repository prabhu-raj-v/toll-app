document.getElementById('doc-file').addEventListener('change', async function (e) {
    const file = e.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch('/parse-doc', {
        method: 'POST',
        body: formData
    });

    const data = await response.json();
    const container = document.getElementById('fields-container');
    container.innerHTML = '';

    if (data.highlights && data.highlights.length > 0) {
        data.highlights.forEach((item, index) => {
            const group = document.createElement('div');
            group.className = 'form-group';

            const label = document.createElement('label');
            label.innerText = `Field ${index + 1}:`;

            const input = document.createElement('input');
            input.type = 'text';
            
            const cleanText = item.text.trim();
            input.name = cleanText;

            // Highlight styling
            const highlightColor = item.color || 'yellow';
            input.style.backgroundColor = highlightColor === 'yellow' ? '#ffffff' : highlightColor;
            input.style.fontWeight = 'bold';
            input.style.color = '#000000';

            // Retrieve from LocalStorage
            const savedValue = localStorage.getItem(`field_pos_${index}`);
            input.value = savedValue !== null ? savedValue : cleanText;

            input.addEventListener('input', (evt) => {
                localStorage.setItem(`field_pos_${index}`, evt.target.value);
            });

            group.appendChild(label);
            group.appendChild(input);
            container.appendChild(group);
        });

        document.getElementById('edit-form').style.display = 'block';
    } else {
        alert('No highlighted text found in this document.');
    }
});

document.getElementById('edit-form').addEventListener('submit', async function (e) {
    e.preventDefault();

    const fileInput = document.getElementById('doc-file');
    const formData = new FormData();
    formData.append('file', fileInput.files[0]);

    const inputs = document.querySelectorAll('#fields-container input');
    inputs.forEach(input => {
        formData.append(input.name, input.value);
    });

    // Get custom filename input value
    let filename = document.getElementById('custom-filename').value.trim();
    if (!filename) {
        filename = 'ripped_document_01';
    }
    if (!filename.endsWith('.docx')) {
        filename += '.docx';
    }

    const response = await fetch('/generate-doc', {
        method: 'POST',
        body: formData
    });

    if (response.ok) {
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        // Assign the user-entered custom filename
        a.download = filename;
        
        document.body.appendChild(a);
        a.click();
        a.remove();
    } else {
        alert('Failed to generate document.');
    }
});