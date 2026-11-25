



 // ============================================================
// ARQUIVO: Sistema.gs
// ============================================================


const SISTEMA = {
  nome: "Nexus",
  versao: "0.9",
  mosteiro: "Mosteiro da Santa Cruz"
};

const CAMPOS = {

  status: "Status",

  nomeCivil: "Nome civil completo",

  nomeReligioso: "Nome religioso",

  chegada: "Data prevista de chegada:",

  partida: "Data prevista de partida:",

  motivo: "Motivo principal da estada:"

};

function sobreSistema() {

  SpreadsheetApp.getUi().alert(

    SISTEMA.nome +
    "\\nVersão " + SISTEMA.versao +

    "\\n\\n" +

    SISTEMA.mosteiro +

    "\\n\\nSistema de apoio à Hospedaria."

  );

}

// ============================================================
// ARQUIVO: Principal.gs
// ============================================================


function sobreSistema() {

  SpreadsheetApp.getUi().alert(

    SISTEMA.nome +
    "\\nVersão " + SISTEMA.versao +

    "\\n\\n" +

    SISTEMA.mosteiro +

    "\\n\\nSistema de apoio à Hospedaria."

  );

}


function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("Alce")

    .addItem("Chegadas", "mostrarTelaChegadas")

    .addItem("📤 Gerar Lista de Saídas", "emDesenvolvimento")

    .addItem("🏠 Hóspedes Presentes", "emDesenvolvimento")

    .addSeparator()

    .addItem("📊 Estatísticas", "emDesenvolvimento")

    .addItem("⚙️ Configurações", "emDesenvolvimento")

    .addSeparator()

    .addItem("ℹ Sobre o Sistema", "sobreSistema")

    .addToUi();

}


function emDesenvolvimento() {

  SpreadsheetApp.getUi().alert(

    "Esta função ainda está em desenvolvimento."

  );

}function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu("Alce")
    .addItem("Chegadas", "mostrarTelaChegadas")
    .addItem("📤 Gerar Lista de Saídas", "emDesenvolvimento")
    .addItem("🏠 Hóspedes Presentes", "emDesenvolvimento")
    .addSeparator()
    .addItem("📊 Estatísticas", "emDesenvolvimento")
    .addItem("⚙️ Configurações", "emDesenvolvimento")
    .addSeparator()
    .addItem("ℹ Sobre o Sistema", "sobreSistema")
    .addToUi();

}

function emDesenvolvimento() {

  SpreadsheetApp.getUi().alert(
    "Esta função ainda está em desenvolvimento."
  );

}

// ============================================================
// ARQUIVO: Operações.gs
// ============================================================


// ============================================================
// LISTA DE CHEGADAS
// ============================================================
//
// O núcleo da operação é:
//
//     obterListaChegadas(mes)
//
// Essa função:
// 1. Lê as duas abas.
// 2. Cria uma cópia atualizada de cada registro.
// 3. Examina TODAS as colunas de modificação.
// 4. Aplica as modificações à cópia.
// 5. Só então verifica Status, CPF e datas.
// 6. Filtra pelo mês solicitado.
// 7. Ordena por data de chegada.
// 8. Devolve a lista.
//
// A função listaChegadas() usa esse núcleo para abrir a interface
// dentro da planilha.
//
// O Web App também usará o mesmo núcleo através de doGet().
//
// As respostas originais dos formulários NUNCA são alteradas.
//
// Regra das modificações:
//
//     X. texto
//
// significa:
//
//     substituir completamente o conteúdo da coluna X
//     pelo texto que aparece depois do ponto.
//
// Exemplo:
//
//     H. 15/10/2026 08:00
//
// faz a coluna H passar a conter a data 15/10/2026 08:00
// na cópia utilizada pelo ALCE.
//
// IMPORTANTE:
//
// As modificações são processadas ANTES da filtragem.
// ============================================================



// ============================================================
// INTERFACE DA PLANILHA
// ============================================================
//
// Mantém o comportamento atual.
//
// Solicita o mês, obtém a lista pelo núcleo e abre
// a interface FolhaChegadas.
//
// ============================================================

function listaChegadas() {

  const ui =
    SpreadsheetApp.getUi();


  // ==========================================================
  // SOLICITA O MÊS
  // ==========================================================

  const mes =
    solicitarMes();

  if (mes === null) {
    return;
  }


  // ==========================================================
  // OBTÉM A LISTA
  // ==========================================================

  const lista =
    obterListaChegadas(
      mes
    );


  // ==========================================================
  // ABRE A FOLHA A4 EM HTML
  // ==========================================================

  const template =
    HtmlService.createTemplateFromFile(
      "FolhaChegadas"
    );


  template.lista =
    lista;


  template.mes =
    mes;


  template.nomeMes =
    obterNomeMes(
      mes
    );


  const html =
    template
      .evaluate()
      .setWidth(900)
      .setHeight(700);


  ui.showModalDialog(
    html,
    "Lista de Chegadas"
  );

}



// ============================================================
// NÚCLEO DA LISTA DE CHEGADAS
// ============================================================
//
// Esta é a parte central da operação.
//
// Ela NÃO abre interface.
//
// Ela NÃO pergunta o mês.
//
// Ela simplesmente recebe o mês e devolve os dados.
//
// Isso permite que:
// - a planilha use a função;
// - o Web App use a mesma função;
// - futuras interfaces usem o mesmo núcleo.
//
// ============================================================

function obterListaChegadas(
  mes
) {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();


  // ==========================================================
  // ABAS
  // ==========================================================

  const abaCadastro =
    planilha.getSheetByName(
      "Cadastro de Hóspedes"
    );

  const abaSolicitacoes =
    planilha.getSheetByName(
      "Solicitação de Hospedagem"
    );


  if (
    !abaCadastro ||
    !abaSolicitacoes
  ) {

    throw new Error(
      "Não foi possível localizar uma das abas necessárias."
    );

  }


  // ==========================================================
  // LEITURA DOS DADOS
  // ==========================================================

  const dadosCadastro =
    abaCadastro
      .getDataRange()
      .getValues();

  const dadosSolicitacoes =
    abaSolicitacoes
      .getDataRange()
      .getValues();


  if (
    dadosCadastro.length === 0 ||
    dadosSolicitacoes.length === 0
  ) {

    throw new Error(
      "Uma das abas não possui dados."
    );

  }


  // ==========================================================
  // CABEÇALHOS E COLUNAS
  // ==========================================================

  const cabecalhoCadastro =
    dadosCadastro[0];

  const cabecalhoSolicitacoes =
    dadosSolicitacoes[0];


  const colCadastro =
    localizarColunas(
      cabecalhoCadastro
    );


  const colSolicitacoes =
    localizarColunas(
      cabecalhoSolicitacoes
    );


  // ==========================================================
  // ÍNDICES — CADASTRO
  // ==========================================================

  const indiceStatusCadastro =
    colCadastro[
      campo("Status")
    ];


  const indiceCPFCadastro =
    colCadastro[
      campo("CPF")
    ];


  const indiceChegadaCadastro =
    colCadastro[
      campo("Previsão de Chegada")
    ];


  const indicePartidaCadastro =
    colCadastro[
      campo("Previsão de Partida")
    ];


  const indiceMotivoCadastro =
    colCadastro[
      campo("Motivo Principal da Hospedagem")
    ];

 const indiceNomeSolicitacao =
  colSolicitacoes[
    campo("Nome")
  ];

  const indiceNomeReligiosoCadastro =
    colCadastro[
      campo("Nome religioso")
    ];


  const indiceNomeCivilCadastro =
    colCadastro[
      campo("Nome civil completo")
    ];


  // ==========================================================
  // ESTADO — CADASTRO DE HÓSPEDES
  //
  // K = coluna 11 da planilha
  // Índice JavaScript = 10
  // ==========================================================

  const indiceEstadoCadastro = 10;



  // ==========================================================
  // ÍNDICES — SOLICITAÇÕES
  // ==========================================================

  const indiceStatusSolicitacao =
    colSolicitacoes[
      campo("Status")
    ];


  const indiceCPFSolicitacao =
    colSolicitacoes[
      campo("CPF")
    ];


  const indiceChegadaSolicitacao =
    colSolicitacoes[
      campo("Previsão de Chegada")
    ];


  const indicePartidaSolicitacao =
    colSolicitacoes[
      campo("Previsão de Partida")
    ];


  const indiceMotivoSolicitacao =
    colSolicitacoes[
      campo("Motivo Principal da Hospedagem")
    ];



  // ==========================================================
  // LISTA FINAL
  // ==========================================================

  const lista = [];



  // ==========================================================
  // 1. PRIMEIRAS HOSPEDAGENS
  //    ABA: CADASTRO DE HÓSPEDES
  // ==========================================================

  if (
    indiceStatusCadastro !== undefined &&
    indiceChegadaCadastro !== undefined &&
    indicePartidaCadastro !== undefined &&
    indiceMotivoCadastro !== undefined
  ) {

    for (
      let i = 1;
      i < dadosCadastro.length;
      i++
    ) {


      // ======================================================
      // ATUALIZA O REGISTRO PRIMEIRO
      // ======================================================

      const registro =
        criarRegistroAtualizado(
          dadosCadastro[i],
          indiceStatusCadastro
        );


      if (!registro) {
        continue;
      }


      const linha =
        registro.linha;



      // ======================================================
      // STATUS ATUALIZADO
      // ======================================================

      const status =
        linha[
          indiceStatusCadastro
        ];


      if (
        String(status)
          .trim()
          .toLowerCase() !== "confirmado"
      ) {

        continue;

      }



      // ======================================================
      // CHEGADA ATUALIZADA
      // ======================================================

      const chegada =
        linha[
          indiceChegadaCadastro
        ];


      if (
        !(chegada instanceof Date) ||
        isNaN(chegada.getTime())
      ) {

        continue;

      }



      // ======================================================
      // FILTRO DO MÊS
      // ======================================================

      if (
        chegada.getMonth() + 1 !== mes
      ) {

        continue;

      }



      // ======================================================
      // NOME ATUALIZADO
      // ======================================================

      const nomeReligioso =
        indiceNomeReligiosoCadastro !== undefined
          ? linha[
              indiceNomeReligiosoCadastro
            ]
          : "";


      const nomeCivil =
        indiceNomeCivilCadastro !== undefined
          ? linha[
              indiceNomeCivilCadastro
            ]
          : "";



      // ======================================================
      // ESTADO ATUALIZADO
      // ======================================================

      const estado =
        linha[
          indiceEstadoCadastro
        ] ?? "";



      // ======================================================
      // ADICIONA À LISTA
      // ======================================================

      lista.push({

        nome:
          nomeReligioso ||
          nomeCivil,

        chegada:
          chegada,

        partida:
          linha[
            indicePartidaCadastro
          ],

        estado:
          estado,

        motivo:
          linha[
            indiceMotivoCadastro
          ]

      });

    }

  }



  // ==========================================================
  // 2. HOSPEDAGENS POSTERIORES
  //    ABA: SOLICITAÇÃO DE HOSPEDAGEM
  // ==========================================================

  if (
  indiceStatusSolicitacao !== undefined &&
  indiceChegadaSolicitacao !== undefined &&
  indicePartidaSolicitacao !== undefined &&
  indiceMotivoSolicitacao !== undefined
) {

    for (
      let i = 1;
      i < dadosSolicitacoes.length;
      i++
    ) {


      // ======================================================
      // ATUALIZA A SOLICITAÇÃO PRIMEIRO
      // ======================================================

      const registroSolicitacao =
        criarRegistroAtualizado(
          dadosSolicitacoes[i],
          indiceStatusSolicitacao
        );


      if (!registroSolicitacao) {
        continue;
      }


      const linhaSolicitacao =
        registroSolicitacao.linha;



      // ======================================================
      // STATUS ATUALIZADO
      // ======================================================

      const status =
        linhaSolicitacao[
          indiceStatusSolicitacao
        ];


      if (
        String(status)
          .trim()
          .toLowerCase() !== "confirmado"
      ) {

        continue;

      }



      // ======================================================
      // CPF ATUALIZADO
      // ======================================================

      const cpf =
        String(
          linhaSolicitacao[
            indiceCPFSolicitacao
          ] ?? ""
        ).trim();



      // ======================================================
      // PROCURA O CPF NO CADASTRO
      // ======================================================

      let cadastro = null;


      for (
        let j = 1;
        j < dadosCadastro.length;
        j++
      ) {


        const registroCadastro =
          criarRegistroAtualizado(
            dadosCadastro[j],
            indiceStatusCadastro
          );


        if (!registroCadastro) {
          continue;
        }


        const linhaCadastro =
          registroCadastro.linha;


        const cpfCadastro =
          String(
            linhaCadastro[
              indiceCPFCadastro
            ] ?? ""
          ).trim();


        if (
          cpf &&
          cpfCadastro === cpf
          ) {

          cadastro =
            registroCadastro;

          break;

        }

      }



      // ======================================================
      // CHEGADA ATUALIZADA DA SOLICITAÇÃO
      // ======================================================

      const chegada =
        linhaSolicitacao[
          indiceChegadaSolicitacao
        ];


      if (
        !(chegada instanceof Date) ||
        isNaN(chegada.getTime())
      ) {

        continue;

      }



      // ======================================================
      // FILTRO DO MÊS
      // ======================================================

      if (
        chegada.getMonth() + 1 !== mes
      ) {

        continue;

      }



     // ======================================================
// NOME E ESTADO
//
// Quando o CPF encontra cadastro, usam-se os dados do
// cadastro atualizado. Caso contrário, usa-se o nome da
// própria solicitação atualizada.
// ======================================================

let nome = "";

let estado = "";


if (cadastro) {

  const linhaCadastro =
    cadastro.linha;


  const nomeReligioso =
    indiceNomeReligiosoCadastro !== undefined
      ? linhaCadastro[
          indiceNomeReligiosoCadastro
        ]
      : "";


  const nomeCivil =
    indiceNomeCivilCadastro !== undefined
      ? linhaCadastro[
          indiceNomeCivilCadastro
        ]
      : "";


  nome =
    nomeReligioso ||
    nomeCivil;


  estado =
    linhaCadastro[
      indiceEstadoCadastro
    ] ?? "";

} else {

  nome =
    indiceNomeSolicitacao !== undefined
      ? linhaSolicitacao[
          indiceNomeSolicitacao
        ]
      : "";

}


if (!String(nome ?? "").trim()) {
  continue;
}



      // ======================================================
      // ADICIONA À LISTA
      // ======================================================

      lista.push({

          nome:
           nome,

        chegada:
          chegada,

        partida:
          linhaSolicitacao[
            indicePartidaSolicitacao
          ],

        estado:
          estado,

        motivo:
          linhaSolicitacao[
            indiceMotivoSolicitacao
          ]

      });

    }

  }



  // ==========================================================
  // 3. ORDENAÇÃO
  // ==========================================================

  lista.sort(
    function(a, b) {

      return (
        a.chegada -
        b.chegada
      );

    }
  );


  return lista;

}


// ============================================================
// CRIA O ESTADO ATUALIZADO DE UM REGISTRO
// ============================================================
//
// Recebe uma linha original e cria uma cópia.
//
// TODAS as colunas posteriores à coluna Status são examinadas.
//
// Se uma célula contiver:
//
//     X. texto
//
// o conteúdo da coluna X da cópia será substituído.
//
// A linha original nunca é modificada.
//
// ============================================================

function criarRegistroAtualizado(
  linhaOriginal,
  indiceStatus
) {

  if (
    !Array.isArray(linhaOriginal) ||
    indiceStatus === undefined
  ) {

    return null;

  }


  // ==========================================================
  // CÓPIA DA LINHA ORIGINAL
  // ==========================================================

  const linha =
    linhaOriginal.slice();



  // ==========================================================
  // EXAMINA TODAS AS COLUNAS DE MODIFICAÇÃO
  // ==========================================================

  for (
    let indice = indiceStatus + 1;
    indice < linhaOriginal.length;
    indice++
  ) {


    const valor =
      String(
        linhaOriginal[indice] ?? ""
      ).trim();


    if (!valor) {
      continue;
    }



    // ========================================================
    // RECONHECE:
    //
    // X. texto
    // AA. texto
    // AB. texto
    // etc.
    // ========================================================

    const correspondencia =
      valor.match(
        /^([A-Z]+)\\.\\s*(.*)$/i
      );


    if (!correspondencia) {
      continue;
    }



    const letra =
      correspondencia[1]
        .toUpperCase();


    const texto =
      correspondencia[2].trim();


    const numeroColuna =
      converterLetraParaNumero(
        letra
      );


    …

}

function registrarFeedback(comentario) {

  if (!comentario || !String(comentario).trim()) {
    throw new Error("O comentário não pode estar vazio.");
  }

  const planilha =
    SpreadsheetApp.openById(
      "190ITb13hbCJuGnoA732I4Cnt3X3TV49PjBqUVtmEacE"
    );

  const aba =
    planilha.getSheetByName("Feedback");

  if (!aba) {
    throw new Error(
      'A aba "Feedback" não foi encontrada.'
    );
  }

  aba.appendRow([
    new Date(),
    String(comentario).trim()
  ]);

  return {
    sucesso: true
  };
}

// ============================================================
// ARQUIVO: Utilitários.gs
// ============================================================


function formatarData(data) {

  if (!(data instanceof Date)) {
    return "";
  }

  return Utilities.formatDate(
    data,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );

}

function localizarColunas(cabecalho) {

  const colunas = {};

  cabecalho.forEach((titulo, indice) => {

    const chave = String(titulo)
      .trim()
      .replace(/\\s+/g, " ");

    colunas[chave] = indice;

  });

  return colunas;

}

function campo(nome) {

  return String(nome)
    .trim()
    .replace(/\\s+/g, " ");

}

function solicitarMes() {

  const ui = SpreadsheetApp.getUi();

  const resposta = ui.prompt(
    "Lista de Chegadas",
    "Digite o mês desejado (1 a 12):",
    ui.ButtonSet.OK_CANCEL
  );

  if (resposta.getSelectedButton() !== ui.Button.OK)
    return null;

  const mes = Number(resposta.getResponseText());

  if (isNaN(mes) || mes < 1 || mes > 12) {

    ui.alert("Mês inválido.");

    return null;

  }

  return mes;

}

function obterDados() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  const aba = planilha.getActiveSheet();

  return aba.getDataRange().getValues();

}

function obterCabecalho(dados) {

  return dados[0];

}

function converterLetraParaNumero(letra) {
  letra = String(letra).trim().toUpperCase();
  if (!/^[A-Z]+$/.test(letra)) return null;
  let numero = 0;
  for (let i = 0; i < letra.length; i++) {
    numero = numero * 26 + (letra.charCodeAt(i) - 64);
  }
  return numero;
}

function obterLetraColuna(numero) {
  let letra = "";
  while (numero > 0) {
    const resto = (numero - 1) % 26;
    letra = String.fromCharCode(65 + resto) + letra;
    numero = Math.floor((numero - 1) / 26);
  }
  return letra;
}

function aplicarComandoNaLista(lista, linhaOrigem, colunas, letra, texto) {
  const numeroColuna = converterLetraParaNumero(letra);
  if (numeroColuna === null) return;
  const indice = numeroColuna - 1;
  if (indice < 0 || indice > 3) return;
  const nomeReligioso = colunas[campo("Nome religioso")] !== undefined ? linhaOrigem[colunas[campo("Nome religioso")]] : "";
  const nomeCivil = colunas[campo("Nome civil completo")] !== undefined ? linhaOrigem[colunas[campo("Nome civil completo")]] : "";
  const nome = nomeReligioso || nomeCivil;
  const chegada = colunas[campo("Previsão de Chegada")] !== undefined ? linhaOrigem[colunas[campo("Previsão de Chegada")]] : null;
  for (let i = 0; i < lista.length; i++) {
    const hospede = lista[i];
    if (hospede.nome === nome && (!chegada || hospede.chegada.getTime() === chegada.getTime())) {
      switch (indice) {
        case 0: hospede.nome = texto; break;
        case 1: hospede.chegada = texto; break;
        case 2: hospede.partida = texto; break;
        case 3: hospede.motivo = texto; break;
      }
      break;
    }
  }
}

// ============================================================
// NEXUS — TESTE DE IDENTIDADE
// ============================================================
// Consulta somente o usuário que está acessando o Web App.
// Não consulta o usuário efetivo (conta proprietária da execução).
// ============================================================

function testarIdentidadeNexus() {
  const email = Session.getActiveUser().getEmail();
  return {
    usuarioAtivo: email || "(vazio)"
  };
}
