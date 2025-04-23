document.addEventListener('DOMContentLoaded', function () {
    const searchInput = document.getElementById('carreta');
    const searchResults = document.getElementById('searchResults');
    const resultsList = document.getElementById('resultsList');

    let debounceTimer;

    function debounce(func, delay) {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(func, delay);
    }

    async function fetchData(query) {
        if (!query || query.length < 1) {
            searchResults.style.display = 'none';
            return;
        }

        try {
            const apiUrl = `/api/listar-carretas?carreta=${encodeURIComponent(query)}`;
            const response = await fetch(apiUrl);
            if (!response.ok) throw new Error('Erro na resposta da API');

            const data = await response.json();
            const lista = data.carretas.map((nome, index) => ({ id: index, name: nome }));
            displayResults(lista);
        } catch (error) {
            console.error('Erro ao buscar carretas:', error);
        }
    }

    function displayResults(data) {
        resultsList.innerHTML = '';

        if (data.length === 0) {
            const noResults = document.createElement('div');
            noResults.className = 'search-item';
            noResults.textContent = 'Nenhum resultado encontrado';
            resultsList.appendChild(noResults);
        } else {
            data.forEach(item => {
                const resultItem = document.createElement('div');
                resultItem.className = 'search-item';
                resultItem.textContent = item.name;
                resultItem.dataset.id = item.id;

                resultItem.addEventListener('click', function () {
                    searchInput.value = item.name;
                    searchResults.style.display = 'none';

                    fetch(`/api/descricao-carreta?carreta=${encodeURIComponent(item.name)}`)
                        .then(r => r.json())
                        .then(data => {
                            document.getElementById("desc").value = data.descricao || "Descrição não encontrada";
                        });

                    fetch(`/api/codigo-carreta?carreta=${encodeURIComponent(item.name)}`)
                        .then(r => r.json())
                        .then(data => {
                            document.getElementById("codigo").value = data.codigo || "Descrição não encontrada";
                        });
                });

                resultsList.appendChild(resultItem);
            });
        }

        searchResults.style.display = 'block';
    }

    searchInput.addEventListener('input', function () {
        const query = this.value.trim();
        debounce(() => fetchData(query), 300);
    });

    document.addEventListener('click', function (e) {
        if (!searchInput.contains(e.target) && !searchResults.contains(e.target)) {
            searchResults.style.display = 'none';
        }
    });

    searchInput.addEventListener('focus', function () {
        if (this.value.trim().length > 0) {
            fetchData(this.value.trim());
        }
    });
});