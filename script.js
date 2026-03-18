const url = 'https://humble-orbit-xxp95jxv5g9c5j5-3000.app.github.dev/';

async function exibirRespostaEmCode() {
	const code = document.createElement('code');
	code.style.display = 'block';
	code.style.whiteSpace = 'pre-wrap';
	code.textContent = 'Carregando...';
	document.body.appendChild(code);

	try {
		const response = await fetch(url, { method: 'GET' });

		if (!response.ok) {
			throw new Error(`Falha na requisição: ${response.status} ${response.statusText}`);
		}

		const resultado = await response.text();
		code.textContent = resultado;
	} catch (error) {
		code.textContent = `Erro ao buscar dados: ${error.message}`;
	}
}

exibirRespostaEmCode();
