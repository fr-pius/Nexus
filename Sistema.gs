



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
    "\nVersão " + SISTEMA.versao +

    "\n\n" +

    SISTEMA.mosteiro +

    "\n\nSistema de apoio à Hospedaria."

  );

}

// ============================================================
// ARQUIVO: Principal.gs
// ============================================================


function sobreSistema() {

  SpreadsheetApp.getUi().alert(

    SISTEMA.nome +
    "\nVersão " + SISTEMA.versao +

    "\n\n" +

    SISTEMA.mosteiro +

    "\n\nSistema de apoio à Hospedaria."

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
// - futuras interfaces usem a mesma função.
//
// ============================================================

function obterListaChegadasInterno_(mes) {

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


function obterListaChegadas(mes) {
  exigirPermissaoNexus_("consultar_chegadas");
  return obterListaChegadasInterno_(mes);
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
        /^([A-Z]+)\.\s*(.*)$/i
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


    if (
      numeroColuna === null
    ) {

      continue;

    }


    const indiceColuna =
      numeroColuna - 1;


    // O Status é uma coluna administrativa especial.
    // Ele nunca é alterado por comandos virtuais: sua alteração
    // é feita diretamente pelo menu suspenso da planilha/Nexus.
    if (indiceColuna === indiceStatus) {
      Logger.log("COMANDO IGNORADO: Status só pode ser alterado diretamente.");
      continue;
    }


    // ========================================================
    // COLUNA ALVO EXISTE?
    // ========================================================

    if (
      indiceColuna < 0 ||
      indiceColuna >= linha.length
    ) {

      Logger.log(
        "MODIFICAÇÃO IGNORADA: " +
        "coluna " +
        letra +
        " não existe na linha."
      );

      continue;

    }



    // ========================================================
    // REGISTRA O COMANDO
    // ========================================================

    Logger.log(
      "COMANDO ENCONTRADO: " +
      valor +
      " | coluna-alvo: " +
      letra +
      " | índice: " +
      indiceColuna +
      " | novo valor: " +
      texto
    );



    // ========================================================
    // DETERMINA SE A COLUNA ALVO É UMA DATA
    // ========================================================

    const colunaEhData =
      verificarSeColunaEhData(
        indiceColuna
      );


    if (colunaEhData) {


      const novaData =
        converterTextoParaData(
          texto
        );


      if (novaData) {

        linha[indiceColuna] =
          novaData;


        Logger.log(
          "DATA ATUALIZADA: " +
          letra +
          " = " +
          novaData
        );

      } else {

        linha[indiceColuna] =
          texto;


        Logger.log(
          "ATENÇÃO: não foi possível " +
          "converter para data: " +
          texto
        );

      }

    } else {


      // ======================================================
      // COLUNA NORMAL
      // ======================================================

      linha[indiceColuna] =
        texto;

    }

  }


  return {
    linha:
      linha
  };

}



// ============================================================
// VERIFICA SE UMA COLUNA É UMA COLUNA DE DATA
// ============================================================

function verificarSeColunaEhData(
  indiceColuna
) {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();


  const abas = [

    planilha.getSheetByName(
      "Cadastro de Hóspedes"
    ),

    planilha.getSheetByName(
      "Solicitação de Hospedagem"
    )

  ];


  for (
    let i = 0;
    i < abas.length;
    i++
  ) {

    const aba =
      abas[i];


    if (!aba) {
      continue;
    }


    const dados =
      aba
        .getDataRange()
        .getValues();


    if (
      dados.length === 0 ||
      indiceColuna >= dados[0].length
    ) {

      continue;

    }


    const cabecalho =
      dados[0];


    const titulo =
      String(
        cabecalho[indiceColuna] ?? ""
      )
      .trim();


    const tituloNormalizado =
      campo(titulo);


    if (
      tituloNormalizado ===
        campo("Previsão de Chegada") ||

      tituloNormalizado ===
        campo("Previsão de Partida")
    ) {

      return true;

    }

  }


  return false;

}



// ============================================================
// CONVERTE TEXTO EM DATA
// ============================================================
//
// Formatos aceitos:
//
// DD/MM/AAAA
// DD/MM/AAAA HH:MM
// DD/MM/AAAA HH:MM:SS
//
// ============================================================

function converterTextoParaData(
  texto
) {

  const valor =
    String(texto).trim();


  const correspondencia =
    valor.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
    );


  if (!correspondencia) {
    return null;
  }


  const dia =
    Number(
      correspondencia[1]
    );


  const mes =
    Number(
      correspondencia[2]
    );


  const ano =
    Number(
      correspondencia[3]
    );


  const hora =
    correspondencia[4] !== undefined
      ? Number(
          correspondencia[4]
        )
      : 0;


  const minuto =
    correspondencia[5] !== undefined
      ? Number(
          correspondencia[5]
        )
      : 0;


  const segundo =
    correspondencia[6] !== undefined
      ? Number(
          correspondencia[6]
        )
      : 0;



  // ==========================================================
  // VALIDA HORÁRIO
  // ==========================================================

  if (
    hora < 0 ||
    hora > 23 ||
    minuto < 0 ||
    minuto > 59 ||
    segundo < 0 ||
    segundo > 59
  ) {

    return null;

  }



  // ==========================================================
  // CRIA A DATA
  // ==========================================================

  const data =
    new Date(
      ano,
      mes - 1,
      dia,
      hora,
      minuto,
      segundo
    );



  // ==========================================================
  // CONFIRMA QUE A DATA EXISTE
  // ==========================================================

  if (
    data.getFullYear() !== ano ||
    data.getMonth() !== mes - 1 ||
    data.getDate() !== dia ||
    data.getHours() !== hora ||
    data.getMinutes() !== minuto ||
    data.getSeconds() !== segundo
  ) {

    return null;

  }


  return data;

}



// ============================================================
// CRIA A FOLHA FORMATADA
// ============================================================
//
// Infraestrutura mantida da versão anterior.
//
// ============================================================

function criarFolhaChegadas(
  lista,
  mes
) {

  const planilha =
    SpreadsheetApp.getActiveSpreadsheet();


  const nomeMes =
    obterNomeMes(
      mes
    );


  const nomeAba =
    "Chegadas - " +
    nomeMes;



  // ==========================================================
  // REMOVE FOLHA ANTERIOR
  // ==========================================================

  const antiga =
    planilha.getSheetByName(
      nomeAba
    );


  if (antiga) {

    planilha.deleteSheet(
      antiga
    );

  }



  // ==========================================================
  // CRIA NOVA FOLHA
  // ==========================================================

  const aba =
    planilha.insertSheet(
      nomeAba
    );



  // ==========================================================
  // TÍTULO
  // ==========================================================

  aba.getRange(
    "A1:D1"
  )
  .merge()
  .setValue(
    "LISTA DE CHEGADAS"
  )
  .setFontFamily(
    "Georgia"
  )
  .setFontSize(
    20
  )
  .setFontWeight(
    "bold"
  )
  .setHorizontalAlignment(
    "center"
  )
  .setVerticalAlignment(
    "middle"
  );



  aba.getRange(
    "A2:D2"
  )
  .merge()
  .setValue(
    nomeMes.toUpperCase()
  )
  .setFontFamily(
    "Georgia"
  )
  .setFontSize(
    13
  )
  .setFontWeight(
    "bold"
  )
  .setHorizontalAlignment(
    "center"
  );



  // ==========================================================
  // CABEÇALHO
  // ==========================================================

  aba.getRange(
    "A4:D4"
  )
  .setValues([[
    "Hóspede",
    "Chegada",
    "Partida",
    "Motivo"
  ]])
  .setFontFamily(
    "Georgia"
  )
  .setFontWeight(
    "bold"
  )
  .setHorizontalAlignment(
    "center"
  )
  .setVerticalAlignment(
    "middle"
  );



  // ==========================================================
  // DADOS
  // ==========================================================

  if (
    lista.length > 0
  ) {

    const valores =
      lista.map(
        function(hospede) {

          return [

            hospede.nome || "",

            hospede.chegada || "",

            hospede.partida || "",

            hospede.motivo || ""

          ];

        }
      );


    aba.getRange(
      5,
      1,
      valores.length,
      4
    )
    .setValues(
      valores
    );



    // --------------------------------------------------------
    // FORMATO DAS DATAS
    // --------------------------------------------------------

    aba.getRange(
      5,
      2,
      valores.length,
      2
    )
    .setNumberFormat(
      "dd/MM/yyyy HH:mm"
    );



    // --------------------------------------------------------
    // FONTE E ALINHAMENTO
    // --------------------------------------------------------

    aba.getRange(
      5,
      1,
      valores.length,
      4
    )
    .setFontFamily(
      "Georgia"
    )
    .setVerticalAlignment(
      "middle"
    );


    aba.getRange(
      5,
      2,
      valores.length,
      2
    )
    .setHorizontalAlignment(
      "center"
    );

  }



  // ==========================================================
  // BORDAS
  // ==========================================================

  const ultimaLinha =
    Math.max(
      4,
      lista.length + 4
    );


  aba.getRange(
    4,
    1,
    ultimaLinha - 3,
    4
  )
  .setBorder(
    true,
    true,
    true,
    true,
    true,
    true
  );



  // ==========================================================
  // LARGURAS
  // ==========================================================

  aba.setColumnWidth(
    1,
    230
  );


  aba.setColumnWidth(
    2,
    125
  );


  aba.setColumnWidth(
    3,
    125
  );


  aba.setColumnWidth(
    4,
    330
  );



  // ==========================================================
  // ALTURAS
  // ==========================================================

  aba.setRowHeight(
    1,
    35
  );


  aba.setRowHeight(
    2,
    25
  );


  aba.setRowHeight(
    4,
    30
  );



  // ==========================================================
  // CONFIGURAÇÕES
  // ==========================================================

  aba.setFrozenRows(
    4
  );


  aba.setHiddenGridlines(
    true
  );


  planilha.setActiveSheet(
    aba
  );

}



// ============================================================
// NOME DOS MESES
// ============================================================

function obterNomeMes(
  mes
) {

  const meses = [

    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro"

  ];


  return meses[
    mes - 1
  ];

}



// ============================================================
// WEB APP — PORTA DE ENTRADA
// ============================================================
//
// A interface web poderá chamar:
//
//     /exec?acao=chegadas&mes=9
//
// O Nexus então executará exatamente a mesma lógica usada
// pela interface da planilha.
//
// ============================================================

function doGet(e) {

  try {

    const acao =
      e && e.parameter
        ? e.parameter.acao
        : null;

    // ============================================================
    // MAPA DO MOSTEIRO
    // ============================================================

    if (
      e &&
      e.parameter &&
      e.parameter.pagina === "mapa"
    ) {

      return HtmlService
        .createHtmlOutputFromFile("mapa")
        .setTitle("Mapa do Mosteiro");

    }
    // ============================================================
    // INTERFACE DO ALCE
    // ============================================================

    if (!acao) {

      return HtmlService
        .createHtmlOutputFromFile("Alce2")
        .setTitle("ALCE — Nexus");

    }


    // ============================================================
    // LISTA DE CHEGADAS
    // ============================================================

    if (acao === "chegadas") {

      const mes =
        Number(e.parameter.mes);


      if (
        !Number.isInteger(mes) ||
        mes < 1 ||
        mes > 12
      ) {

        return respostaJson({
          erro:
            "Mês inválido. Informe um valor entre 1 e 12."
        });

      }


      const lista =
        obterListaChegadas(mes);


      const listaSerializada =
  lista.map(function(item) {

    return {

      nome:
        item.nome,

      chegada:
        item.chegada instanceof Date
          ? item.chegada.toISOString()
          : String(
              item.chegada || ""
            ),

      partida:
        item.partida instanceof Date
          ? item.partida.toISOString()
          : String(
              item.partida || ""
            ),

      estado:
        item.estado || "",

      motivo:
        item.motivo || ""

    };

  });


      return respostaJson({

        sistema:
          SISTEMA.nome,

        versao:
          SISTEMA.versao,

        mes:
          mes,

        nomeMes:
          obterNomeMes(mes),

        lista:
          listaSerializada

      });

    }

    if (acao === "folhaChegadas") {

  exigirPermissaoNexus_("consultar_chegadas");

  const token =
    e.parameter
      ? e.parameter.token
      : null;

  if (!token) {

    return HtmlService
      .createHtmlOutput(
        "<p>Não foi possível abrir a folha de chegadas.</p>"
      );

  }

  const cache =
    CacheService.getScriptCache();

  const dadosJson =
    cache.get("folhaChegadas_" + token);

  if (!dadosJson) {

    return HtmlService
      .createHtmlOutput(
        "<p>A folha de chegadas expirou ou não está mais disponível.</p>"
      );

  }

  const dados =
  JSON.parse(dadosJson);


const listaPreparada =
  (dados.lista || []).map(function(item) {

    return {

      nome:
        item.nome || "",

      chegada:
        item.chegada
          ? new Date(item.chegada)
          : "",

      partida:
        item.partida
          ? new Date(item.partida)
          : "",

      estado:
        item.estado || "",

      motivo:
        item.motivo || ""

    };

  });

const template =
  HtmlService
    .createTemplateFromFile(
      "FolhaChegadas"
    );


template.lista =
  listaPreparada;

template.nomeMes =
  dados.nomeMes || "";

  return template
    .evaluate()
    .setTitle("Lista de Chegadas");

}

    // ============================================================
    // AÇÃO DESCONHECIDA
    // ============================================================

    return respostaJson({
      erro: "Ação desconhecida."
    });


  } catch (erro) {

    return respostaJson({
      erro:
        String(
          erro.message || erro
        )
    });

  }

}

// ============================================================
// CONSULTA DO ALCE — INTERFACE WEB
// ============================================================

function consultarChegadasPeloAlce(meses) {

  const listaMeses = Array.isArray(meses)
    ? meses.map(Number)
    : [Number(meses)];

  if (
    !listaMeses.length ||
    listaMeses.some(function(mes) {
      return !Number.isInteger(mes) || mes < 1 || mes > 12;
    })
  ) {
    throw new Error(
      "Seleção inválida. Informe um ou mais meses entre 1 e 12."
    );
  }

  exigirPermissaoNexus_("consultar_chegadas");

  const mesesUnicos = Array.from(new Set(listaMeses));

  let lista = [];

  mesesUnicos.forEach(function(mes) {
    lista = lista.concat(obterListaChegadasInterno_(mes));
  });

  lista.sort(function(a, b) {
    return a.chegada - b.chegada;
  });

  return lista.map(function(item) {

    return {

      nome:
        item.nome,

      chegada:
        item.chegada instanceof Date
          ? item.chegada.toISOString()
          : String(
              item.chegada || ""
            ),

      partida:
        item.partida instanceof Date
          ? item.partida.toISOString()
          : String(
              item.partida || ""
            ),

      estado:
        item.estado || "",

      motivo:
        item.motivo || ""

    };

  });

}

function prepararFolhaChegadas(lista, nomeMes) {

  exigirPermissaoNexus_("consultar_chegadas");

  const token = Utilities.getUuid();

  const dados = {
    lista: lista || [],
    nomeMes: nomeMes || ""
  };

  

  CacheService
    .getScriptCache()
    .put(
      "folhaChegadas_" + token,
      JSON.stringify(dados),
      600
    );

  const urlBase =
    ScriptApp.getService().getUrl();

  return (
    urlBase +
    "?acao=folhaChegadas&token=" +
    encodeURIComponent(token)
  );
}

// ============================================================
// RESPOSTA JSON
// ============================================================

function respostaJson(
  dados
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(
        dados
      )
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

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
      .replace(/\s+/g, " ");

    colunas[chave] = indice;

  });

  return colunas;

}

function campo(nome) {

  return String(nome)
    .trim()
    .replace(/\s+/g, " ");

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
// ================================================================
// CONVERTE LETRA DA COLUNA EM NÚMERO
// ================================================================
//
// A  = 1
// Z  = 26
// AA = 27
// AB = 28
//
// ================================================================

function converterLetraParaNumero(letra) {

  letra =
    String(letra)
      .trim()
      .toUpperCase();

  if (!/^[A-Z]+$/.test(letra)) {
    return null;
  }

  let numero = 0;

  for (
    let i = 0;
    i < letra.length;
    i++
  ) {

    numero =
      numero * 26 +
      (
        letra.charCodeAt(i) - 64
      );

  }

  return numero;

}
// ================================================================
// CONVERTE NÚMERO DA COLUNA EM LETRA
// ================================================================

function obterLetraColuna(numero) {

  let letra = "";

  while (numero > 0) {

    const resto =
      (numero - 1) % 26;

    letra =
      String.fromCharCode(65 + resto) +
      letra;

    numero =
      Math.floor((numero - 1) / 26);

  }

  return letra;

}


// ================================================================
// APLICA UM COMANDO À LISTA GERADA
// ================================================================

function aplicarComandoNaLista(
  lista,
  linhaOrigem,
  colunas,
  letra,
  texto
) {

  /*
   * Por enquanto o comando é interpretado
   * de acordo com os campos que efetivamente
   * aparecem na lista.
   *
   * L → nome
   * M → chegada
   * N → partida
   * O → motivo
   *
   * Isso nos permite testar a infraestrutura
   * sem alterar o banco de dados.
   */

  const numeroColuna =
    converterLetraParaNumero(letra);

  if (numeroColuna === null) {
    return;
  }


  // --------------------------------------------------------------
  // ATENÇÃO:
  //
  // A lista impressa possui atualmente quatro campos:
  //
  // A = Nome
  // B = Chegada
  // C = Partida
  // D = Motivo
  //
  // Portanto, nesta primeira implementação,
  // só aplicamos modificações a esses campos.
  // --------------------------------------------------------------

  const indice =
    numeroColuna - 1;

  if (indice < 0 || indice > 3) {
    return;
  }


  // --------------------------------------------------------------
  // Identifica a hospedagem correspondente.
  //
  // O vínculo é feito pelos dados de origem,
  // sem modificar a planilha.
  // --------------------------------------------------------------

  const nomeReligioso =
    colunas[campo("Nome religioso")] !== undefined
      ? linhaOrigem[
          colunas[campo("Nome religioso")]
        ]
      : "";

  const nomeCivil =
    colunas[campo("Nome civil completo")] !== undefined
      ? linhaOrigem[
          colunas[campo("Nome civil completo")]
        ]
      : "";

  const nome =
    nomeReligioso || nomeCivil;

  const chegada =
    colunas[campo("Previsão de Chegada")] !== undefined
      ? linhaOrigem[
          colunas[campo("Previsão de Chegada")]
        ]
      : null;


  // --------------------------------------------------------------
  // Procura a pessoa correspondente na lista.
  // --------------------------------------------------------------

  for (let i = 0; i < lista.length; i++) {

    const hospede =
      lista[i];

    if (
      hospede.nome === nome &&
      (
        !chegada ||
        hospede.chegada.getTime() ===
        chegada.getTime()
      )
    ) {

      switch (indice) {

        case 0:
          hospede.nome = texto;
          break;

        case 1:
          hospede.chegada = texto;
          break;

        case 2:
          hospede.partida = texto;
          break;

        case 3:
          hospede.motivo = texto;
          break;

      }

      break;

    }

  }

}


// ================================================================
// CONVERTE LETRA DE COLUNA EM NÚMERO
// ================================================================

function converterLetraParaNumero(letra) {

  letra =
    String(letra)
      .trim()
      .toUpperCase();

  if (!/^[A-Z]+$/.test(letra)) {
    return null;
  }

  let numero = 0;

  for (let i = 0; i < letra.length; i++) {

    numero =
      numero * 26 +
      (
        letra.charCodeAt(i) -
        64
      );

  }

  return numero;

}

// ============================================================
// ARQUIVO: Manutenção.gs
// ============================================================


function gerarRecursoImagem() {

  const id = "1b8SzTrTFPaW9TWF9r__anLNhWUgP21U_";

  const blob = DriveApp.getFileById(id).getBlob();

  const base64 =
    Utilities.base64Encode(blob.getBytes());

  const codigo =
`const Recursos = {

  imagemPadrao() {
    return "data:${blob.getContentType()};base64,${base64}";
  }

};`;

  DriveApp.createFile(
    "RecursosGerado.gs",
    codigo,
    MimeType.PLAIN_TEXT
  );

  SpreadsheetApp.getUi().alert(
    "Arquivo RecursosGerado.gs criado no seu Drive."
  );

}

// ============================================================
// ARQUIVO: Interface.gs
// ============================================================


function mostrarTelaChegadas() {

  const template =
    HtmlService.createTemplateFromFile("TelaAlce");

  template.imagemAlce =
    "https://drive.google.com/uc?export=view&id=1fdu7KZqGww3indzRqfPVpguZG0sLfXhm";

  template.titulo = "Chegadas";

  template.mensagem =
    "Bom dia, irmão.<br><br>Vamos trabalhar com as chegadas?";

  const html =
    template
      .evaluate()
      .setWidth(520)
      .setHeight(500);

  SpreadsheetApp.getUi()
    .showModalDialog(html, "ALCE");

}

// ============================================================
// ARQUIVO: Recursos.gs
// ============================================================

// Arquivo atualmente vazio.


// ============================================================
// NEXUS — ADMINISTRAÇÃO DE SOLICITAÇÕES
// ============================================================
// A fila é alimentada pela aba "Solicitação de Hospedagem".
// As respostas do formulário nunca são regravadas. Os comandos
// virtuais são aplicados apenas à cópia lida pelo Nexus.
// O Status é a única célula que esta operação modifica.
// ============================================================

const NEXUS_STATUS_VALIDOS = [
  "Confirmado",
  "Cancelado",
  "Em andamento",
  "Concluído",
  "Neutro"
];

function nexusNormalizarValorAssinatura(valor) {
  if (valor instanceof Date) {
    return { tipo: "data", valor: valor.getTime() };
  }
  if (valor === null || valor === undefined) {
    return { tipo: "nulo", valor: "" };
  }
  return { tipo: typeof valor, valor: valor };
}

function nexusAssinaturaLinha(linha, nomeAba) {
  const conteudo = JSON.stringify({
    aba: String(nomeAba || ""),
    linha: linha.map(nexusNormalizarValorAssinatura)
  });
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    conteudo,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64EncodeWebSafe(digest);
}

function nexusValorParaInterface(valor) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm"
    );
  }
  if (valor === null || valor === undefined) return "";
  return String(valor);
}

/**
 * Devolve solicitações ainda abertas para avaliação.
 * Exibe todos os campos do formulário (antes da coluna Status),
 * mas nunca exibe as colunas que contêm comandos de modificação.
 */
function obterSolicitacoesPendentesNexus() {
  exigirPermissaoNexus_("consultar_solicitacoes");

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const nomesAbas = [
    "Cadastro de Hóspedes",
    "Solicitação de Hospedagem"
  ];
  const pendentes = [];

  nomesAbas.forEach(function(nomeAba) {
    const aba = planilha.getSheetByName(nomeAba);
    if (!aba) {
      throw new Error('A aba "' + nomeAba + '" não foi encontrada.');
    }

    const dados = aba.getDataRange().getValues();
    if (!dados || dados.length < 2) return;

    const cabecalho = dados[0];
    const colunas = localizarColunas(cabecalho);
    const indiceStatus = colunas[campo("Status")];
    if (indiceStatus === undefined) {
      throw new Error('A coluna "Status" não foi encontrada na aba "' + nomeAba + '".');
    }

    for (let i = 1; i < dados.length; i++) {
      const linhaOriginal = dados[i];

      // A primeira coluna contém o carimbo de data e hora.
      // Sem esse valor, a linha não representa uma resposta de formulário.
      const carimbo = linhaOriginal[0];
      if (carimbo === "" || carimbo === null || carimbo === undefined ||
          String(carimbo).trim() === "") {
        continue;
      }

      // Somente registros cujo Status esteja realmente vazio aguardam avaliação.
      const statusOriginal = linhaOriginal[indiceStatus];
      if (statusOriginal !== "" && statusOriginal !== null &&
          statusOriginal !== undefined && String(statusOriginal).trim() !== "") {
        continue;
      }

      // Aplicam-se os comandos virtuais apenas à cópia lida pelo Nexus.
      const registro = criarRegistroAtualizado(linhaOriginal, indiceStatus);
      if (!registro) continue;

      const campos = [];
      for (let c = 0; c < indiceStatus; c++) {
        const titulo = String(cabecalho[c] || "").trim();
        if (!titulo) continue;
        campos.push({
          nome: titulo,
          valor: nexusValorParaInterface(registro.linha[c])
        });
      }

      pendentes.push({
        assinatura: nexusAssinaturaLinha(linhaOriginal, nomeAba),
        origem: nomeAba,
        status: "",
        campos: campos
      });
    }
  });

  return pendentes;
}

/**
 * Atualiza exclusivamente a célula Status da solicitação indicada.
 * A assinatura da linha é conferida novamente sob bloqueio; assim,
 * se a linha foi alterada enquanto a tela estava aberta, a gravação
 * é recusada em vez de atingir outro hóspede.
 */
function atualizarStatusSolicitacaoNexus(assinatura, novoStatus) {
  exigirPermissaoNexus_("atualizar_status");

  const status = String(novoStatus || "").trim();
  if (!NEXUS_STATUS_VALIDOS.includes(status)) {
    throw new Error("Status inválido. Escolha uma das opções disponíveis.");
  }
  if (!assinatura) {
    throw new Error("Não foi possível identificar o registro. Atualize a lista.");
  }

  const bloqueio = LockService.getScriptLock();
  bloqueio.waitLock(10000);
  try {
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const nomesAbas = [
      "Cadastro de Hóspedes",
      "Solicitação de Hospedagem"
    ];
    const correspondencias = [];

    nomesAbas.forEach(function(nomeAba) {
      const aba = planilha.getSheetByName(nomeAba);
      if (!aba) {
        throw new Error('A aba "' + nomeAba + '" não foi encontrada.');
      }

      const dados = aba.getDataRange().getValues();
      if (!dados || dados.length < 2) return;

      const colunas = localizarColunas(dados[0]);
      const indiceStatus = colunas[campo("Status")];
      if (indiceStatus === undefined) {
        throw new Error('A coluna "Status" não foi encontrada na aba "' + nomeAba + '".');
      }

      for (let i = 1; i < dados.length; i++) {
        const linha = dados[i];
        const carimbo = linha[0];
        if (carimbo === "" || carimbo === null || carimbo === undefined ||
            String(carimbo).trim() === "") {
          continue;
        }

        if (nexusAssinaturaLinha(linha, nomeAba) === assinatura) {
          correspondencias.push({
            aba: aba,
            nomeAba: nomeAba,
            numeroLinha: i + 1,
            indiceStatus: indiceStatus,
            statusAtual: linha[indiceStatus]
          });
        }
      }
    });

    if (correspondencias.length !== 1) {
      throw new Error(
        correspondencias.length === 0
          ? "O registro foi modificado ou movido desde a consulta. Atualize a lista antes de tentar novamente."
          : "Foram encontrados registros idênticos. Por segurança, o Nexus não alterou nenhum deles."
      );
    }

    const alvo = correspondencias[0];
    if (alvo.statusAtual !== "" && alvo.statusAtual !== null &&
        alvo.statusAtual !== undefined && String(alvo.statusAtual).trim() !== "") {
      throw new Error("Este registro já recebeu um Status. Atualize a lista.");
    }

    // ÚNICA escrita desta operação: uma célula na coluna Status da aba de origem.
    alvo.aba.getRange(alvo.numeroLinha, alvo.indiceStatus + 1).setValue(status);
    SpreadsheetApp.flush();

    return {
      sucesso: true,
      status: status,
      origem: alvo.nomeAba
    };
  } finally {
    bloqueio.releaseLock();
  }
}

/**
 * Verifica a autorização no cadastro central de acessos.
 * A identidade vem exclusivamente da sessão Google no servidor.
 * O sufixo "_" impede a chamada direta pelo cliente via google.script.run.
 */
function exigirPermissaoNexus_(operacao) {
  const email = String(
    Session.getActiveUser().getEmail() || ""
  ).trim().toLowerCase();

  if (!email) {
    throw new Error(
      "Acesso negado: não foi possível identificar sua conta Google."
    );
  }

  const planilhaAcessos = SpreadsheetApp.openById(
    "1tT8PRv3VBEjL7Jp_1dUcDPqQ3BtdplYLigG3bT2DZL8"
  );

  const aba = planilhaAcessos.getSheetByName("Permissoes");

  if (!aba) {
    throw new Error(
      "O controle de acessos está indisponível. Contate o responsável pelo Nexus."
    );
  }

  const ultimaColuna = aba.getLastColumn();
  const ultimaLinha = aba.getLastRow();

  if (ultimaColuna < 1 || ultimaLinha < 2) {
    throw new Error(
      "Acesso negado: sua conta não está autorizada para esta operação."
    );
  }

  const cabecalhos = aba
    .getRange(1, 1, 1, ultimaColuna)
    .getDisplayValues()[0];

  const indiceOperacao = cabecalhos.findIndex(function(cabecalho) {
    return String(cabecalho || "").trim() === operacao;
  });

  if (indiceOperacao < 0) {
    throw new Error(
      "Acesso negado: sua conta não está autorizada para esta operação."
    );
  }

  const emailsAutorizados = aba
    .getRange(2, indiceOperacao + 1, ultimaLinha - 1, 1)
    .getDisplayValues()
    .map(function(linha) {
      return String(linha[0] || "").trim().toLowerCase();
    });

  if (!emailsAutorizados.includes(email)) {
    throw new Error(
      "Acesso negado: sua conta não está autorizada para esta operação."
    );
  }

  return email;
}




