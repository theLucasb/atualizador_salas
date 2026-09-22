require("dotenv").config();

const XLSX = require("xlsx");
const axios = require("axios");
const fs = require("fs");

const {
  GITHUB_TOKEN,
  GITHUB_OWNER,
  GITHUB_REPO,
  EXCEL_PATH
} = process.env;

function obterNomeAbaAtual() {

  const hoje = new Date();

  const meses = [
    "JANEIRO",
    "FEVEREIRO",
    "MARÇO",
    "ABRIL",
    "MAIO",
    "JUNHO",
    "JULHO",
    "AGOSTO",
    "SETEMBRO",
    "OUTUBRO",
    "NOVEMBRO",
    "DEZEMBRO"
  ];

  return `${meses[hoje.getMonth()]} ${hoje.getFullYear()}`;
}

function excelDateToISO(valor) {

  if (!valor) return null;

  if (typeof valor === "number") {

    const data =
      XLSX.SSF.parse_date_code(valor);

    return `${data.y}-${String(data.m).padStart(2, "0")}-${String(data.d).padStart(2, "0")}`;
  }

  if (valor instanceof Date) {

    return valor
      .toISOString()
      .split("T")[0];
  }

  return null;
}

function gerarDados(workbook) {

  const nomeAba = obterNomeAbaAtual();

  console.log(`Lendo aba: ${nomeAba}`);

  const worksheet =
    workbook.Sheets[nomeAba];

  if (!worksheet) {

    throw new Error(
      `A aba '${nomeAba}' não foi encontrada.`
    );
  }

  const linhas =
    XLSX.utils.sheet_to_json(
      worksheet,
      {
        header: 1,
        defval: ""
      }
    );

  const cabecalhos = linhas[0];

  const dados = [];

  const ambientesPermitidos = [
    "SALA 01",
    "SALA 02",
    "SALA 03",
    "SALA MÓVEL 01",
    "SALA MÓVEL 02",
    "LAB. INFORMÁTICA",
    "LAB. AUTOMAÇÃO",
    "LAB. ELÉTRICA",
    "LAB. MECÂNICA",
    "LAB. VESTUÁRIO"
  ];

  for (let i = 1; i < linhas.length; i++) {

    const linha = linhas[i];

    const diaSemana = linha[0];
    const data = excelDateToISO(linha[1]);

    if (!data) continue;

    for (
      let coluna = 2;
      coluna < cabecalhos.length;
      coluna++
    ) {

      const sala =
        cabecalhos[coluna]
          ?.toString()
          .trim();

      if (
        !ambientesPermitidos.includes(sala)
      ) {
        continue;
      }

      const turma = linha[coluna];

      if (
        turma &&
        turma.toString().trim() !== ""
      ) {

        dados.push({
          data,
          diaSemana: diaSemana
            ?.toString()
            .trim(),

          sala,

          turma: turma
            ?.toString()
            .trim()
        });
      }
    }
  }

  return dados;
}

async function atualizarGithub() {

  console.log("=================================");
  console.log("ATUALIZADOR DE SALAS -", new Date().toLocaleString("pt-BR"));
  console.log("=================================");

  try {

    const stats =
      fs.statSync(EXCEL_PATH);

    console.log(
      "Última alteração do Excel:",
      stats.mtime
    );

    console.log("Abrindo planilha...");

    const workbook =
      XLSX.readFile(EXCEL_PATH);

    const dados =
      gerarDados(workbook);

    console.log(
      `${dados.length} registros encontrados`
    );

    const jsonString =
      JSON.stringify(dados, null, 2);

    console.log(
      "Buscando SHA atual..."
    );

    const arquivoAtual =
      await axios.get(
        `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/dados.json`,
        {
          headers: {
            Authorization:
              `Bearer ${GITHUB_TOKEN}`
          }
        }
      );

    const sha =
      arquivoAtual.data.sha;

    const conteudoAtual =
      Buffer
        .from(
          arquivoAtual.data.content,
          "base64"
        )
        .toString("utf8");

    if (conteudoAtual === jsonString) {

      console.log(
        "Nenhuma alteração encontrada."
      );

      return;
    }

    console.log(
      "Enviando atualização..."
    );

    await axios.put(
      `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/dados.json`,
      {
        message:
          `Atualização automática ${new Date().toLocaleString("pt-BR")}`,

        content:
          Buffer
            .from(jsonString)
            .toString("base64"),

        sha
      },
      {
        headers: {
          Authorization:
            `Bearer ${GITHUB_TOKEN}`
        }
      }
    );

    console.log(
      "Atualização concluída!"
    );

    console.log(
      `Total enviado: ${dados.length} registros`
    );

  } catch (err) {

    console.error("ERRO ao atualizar:");
    console.error(err.response?.data || err.message || err);
  }
}

// ================================================
// NOVO: mantém o processo vivo e observa a planilha
// ================================================

let debounceTimer = null;

function agendarAtualizacao(motivo) {

  console.log(`\n${motivo} — aguardando 3s antes de atualizar...`);

  clearTimeout(debounceTimer);

  debounceTimer = setTimeout(() => {
    atualizarGithub();
  }, 3000); // espera 3s de "silêncio" após a última alteração
}

console.log(`Observando alterações em: ${EXCEL_PATH}`);
console.log("Este processo precisa ficar rodando (use pm2, Task Scheduler ou similar).\n");

// roda uma vez assim que o script inicia
atualizarGithub();

// dispara sempre que o arquivo for salvo
fs.watch(EXCEL_PATH, (eventType) => {

  if (eventType === "change") {
    agendarAtualizacao("Alteração detectada na planilha");
  }
});

// rede de segurança: caso o evento de "change" não seja disparado
// (acontece às vezes com arquivos sincronizados via OneDrive/SharePoint),
// verifica a cada 2 minutos se algo mudou
setInterval(() => {
  agendarAtualizacao("Verificação periódica");
}, 2 * 60 * 1000);