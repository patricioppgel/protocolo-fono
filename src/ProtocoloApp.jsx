import React, { useState, useEffect, useMemo } from "react";
import _ from "lodash";
import * as XLSX from "xlsx";
import {
  Plus, Search, Printer, Pencil, Trash2, Save, X, ChevronDown,
  ChevronRight, ClipboardList, User, Calendar, ArrowLeft, AlertCircle,
  BookOpen, FileBarChart, Ear, Baby, HeartPulse, Mic2, Utensils, Brain,
  MessageCircle, Building2, GraduationCap, FileSpreadsheet, Activity, Smile, Check,
} from "lucide-react";

/* ---------------------------------------------------------------------
   Design tokens — cool clinical grey-sage paper, deep teal accent.
------------------------------------------------------------------------ */
const T = {
  paper: "#EEF2F1", surface: "#FFFFFF", surfaceSoft: "#F5F7F6",
  ink: "#1D2A28", inkSoft: "#54655F", inkFaint: "#8B9B95",
  accent: "#2C6E62", accentDeep: "#1E4E45", accentSoft: "#DCEAE5",
  line: "#D2DBD8", lineSoft: "#E4EAE8", warn: "#A6432F", warnSoft: "#F3DFDA",
};
const FONT_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=IBM+Plex+Sans:wght@400;500;600&display=swap');";

/* ---------------------------------------------------------------------
   Schema DSL helpers
   chk  -> checkbox group, optionally with a trailing free-text field
   txt  -> single-line text field
   ta   -> multi-line text field
   grp  -> named sub-block wrapping several fields (renders boxed)
------------------------------------------------------------------------ */
const chk = (path, options, extra) => ({ type: "checks", path, options, extra });
const txt = (path, label) => ({ type: "text", path, label });
const ta = (path, label, rows = 3) => ({ type: "textarea", path, label, rows });
const grp = (title, children) => ({ type: "group", title, children });
const cst = (path, blank, Edit, Read, textLines) => ({ type: "custom", path, blank, Edit, Read, textLines });
const emptyChecks = (list) => Object.fromEntries(list.map(([k]) => [k, false]));

function blankFromNodes(nodes, acc) {
  nodes.forEach((n) => {
    if (n.type === "group") blankFromNodes(n.children, acc);
    else if (n.type === "checks") {
      _.set(acc, n.path, emptyChecks(n.options));
      if (n.extra) _.set(acc, n.extra.path, "");
    } else if (n.type === "custom") {
      _.set(acc, n.path, n.blank());
    } else {
      _.set(acc, n.path, "");
    }
  });
}
function blankFromSections(sections) {
  const acc = {};
  sections.forEach((s) => blankFromNodes(s.children, acc));
  return acc;
}
function blankClinico(areaId) {
  return blankFromSections(AREAS[areaId].sections);
}

/* ---------------------------------------------------------------------
   Area 1 — Linguagem Infantil, Fala e Motricidade Orofacial
   (mirrors the UNIFAN paper protocol exactly)
------------------------------------------------------------------------ */
const A1_S3 = [["linguagemReceptiva","Linguagem receptiva"],["linguagemExpressiva","Linguagem expressiva"],["pragmatica","Pragmática"],["fala","Fala"],["leitura","Leitura"],["escrita","Escrita"],["motricidadeOrofacial","Motricidade Orofacial"]];
const A1_S4 = [["jogos","Jogos"],["livros","Livros"],["figuras","Figuras"],["sequencias","Sequências"],["espelho","Espelho"],["tablet","Tablet"],["caa","CAA"],["materiaisMO","Materiais de MO"]];
const A1_ENTRADA = [["espontanea","Espontânea"],["mediacao","Mediação"],["choro","Choro"],["irritado","Irritado"],["sonolento","Sonolento"]];
const A1_ATENCAO = [["sustentada","Sustentada"],["flutuante","Flutuante"],["dispersa","Dispersa"]];
const A1_COMPORT = [["colaborativo","Colaborativo"],["opositor","Opositor"],["impulsivo","Impulsivo"],["inflexivel","Inflexível"]];
const A1_COMUN = [["espontanea","Espontânea"],["mediada","Mediada"],["gestual","Gestual"]];
const A1_RECEPT = [["ordensSimples","Ordens simples"],["ordensComplexas","Ordens complexas"]];
const A1_EXPRESS = [["palavra","Palavra"],["frases","Frases"],["narrativa","Narrativa"]];
const A1_PRAGM = [["iniciaDialogo","Inicia diálogo"],["mantemDialogo","Mantém diálogo"],["trocaTurnos","Troca turnos"],["mantemTopico","Mantém tópico"]];
const A1_FALA = [["omissoes","Omissões"],["substituicoes","Substituições"],["distorcoes","Distorções"],["simplificacoes","Simplificações"]];
const A1_LEIT = [["decodificacao","Decodificação"],["fluencia","Fluência"],["compreensao","Compreensão"]];
const A1_ESCR = [["palavras","Palavras"],["frases","Frases"],["texto","Texto"]];
const A1_MO = [["face","Face"],["labios","Lábios"],["lingua","Língua"],["respiracao","Respiração"],["mastigacao","Mastigação"],["degluticao","Deglutição"]];

const AREA_LINGUAGEM_FALA_MO = {
  label: "Linguagem Infantil, Fala e Motricidade Orofacial",
  disciplina: "Estágio Supervisionado I",
  icon: MessageCircle,
  sections: [
    { title: "Objetivos da Sessão", children: [ chk("checks", A1_S3, { path: "outro", label: "Outro" }) ] },
    { title: "Recursos Utilizados", children: [ chk("checks", A1_S4, { path: "outros", label: "Outros" }) ] },
    { title: "Observação Comportamental", children: [
      grp("Entrada", [ chk("entrada.checks", A1_ENTRADA), txt("entrada.obs", "Obs.") ]),
      grp("Atenção", [ chk("atencao.checks", A1_ATENCAO), txt("atencao.obs", "Obs.") ]),
      grp("Comportamento", [ chk("comportamento.checks", A1_COMPORT) ]),
      ta("obsGeral", "Observações", 2),
    ]},
    { title: "Linguagem", children: [
      grp("Comunicação", [ chk("comunicacao.checks", A1_COMUN), ta("comunicacao.obs", "Obs.", 2) ]),
      grp("Receptiva", [ chk("receptiva.checks", A1_RECEPT), ta("receptiva.obs", "Obs.", 2) ]),
      grp("Expressiva", [ chk("expressiva.checks", A1_EXPRESS), ta("expressiva.obs", "Obs.", 2) ]),
      grp("Pragmática", [ chk("pragmatica.checks", A1_PRAGM), ta("pragmatica.obs", "Observações", 2) ]),
    ]},
    { title: "Fala", children: [ chk("checks", A1_FALA), ta("fonemas", "Fonemas / processos observados", 3) ] },
    { title: "Leitura e Escrita", children: [
      grp("Leitura", [ chk("leitura", A1_LEIT) ]),
      grp("Escrita", [ chk("escrita", A1_ESCR) ]),
      ta("obs", "Observações", 2),
    ]},
    { title: "Motricidade Orofacial", children: [ chk("checks", A1_MO), ta("obs", "Observações", 3) ] },
  ],
};

/* ---------------------------------------------------------------------
   Area 2 — Audiologia (Adulto e Infantil)
   Grounded in the official UNIFAN forms provided: Anamnese Audiológica —
   Adulto, Anamnese Audiológica Infantil, and Folha de Registro de
   Audiometria (audiograma, logoaudiometria, imitanciometria).
------------------------------------------------------------------------ */
const FREQS_AUD = [125, 250, 500, 750, 1000, 1500, 2000, 3000, 4000, 6000, 8000];
const REFLEX_FREQS = ["500", "1000", "2000", "4000"];
const blankEarThresholds = () => ({ va: Object.fromEntries(FREQS_AUD.map((f) => [f, ""])), vo: Object.fromEntries(FREQS_AUD.map((f) => [f, ""])) });
const blankAudiograma = () => ({ od: blankEarThresholds(), oe: blankEarThresholds(), laudo: { od: "", oe: "", obs: "", referencia: "" }, mascaramentos: [] });
const blankLogoaudio = () => ({
  od: { srt: "", ldf: "", masc: "" }, oe: { srt: "", ldf: "", masc: "" },
  iprf: { od: { intensidade: "", mono: "", dissi: "", masc: "" }, oe: { intensidade: "", mono: "", dissi: "", masc: "" } },
});
const blankTimpanometria = () => ({
  od: { pressao: "", complacencia: "", volOE: "", gradiente: "", largura: "" },
  oe: { pressao: "", complacencia: "", volOE: "", gradiente: "", largura: "" },
  classificacao: { od: "", oe: "", obs: "", referencia: "" },
});
const blankReflexos = () => ({
  od: Object.fromEntries(REFLEX_FREQS.map((f) => [f, { limiar: "", racl: "", dif: "", rail: "" }])),
  oe: Object.fromEntries(REFLEX_FREQS.map((f) => [f, { limiar: "", racl: "", dif: "", rail: "" }])),
  classificacao: { od: "", oe: "", obs: "", referencia: "" },
});

const AUD_QUEIXA_ADULTO = [["dificuldadeParaOuvir","Dificuldade para ouvir"],["zumbido","Zumbido"],["tonturaVertigem","Tontura/vertigem"],["ouvidoTampado","Ouvido tampado"],["dorDeOuvido","Dor de ouvido"],["secrecao","Secreção"],["perdaAuditivaSubita","Perda auditiva súbita"],["acompanhamento","Acompanhamento"],["avaliacaoOcupacional","Avaliação ocupacional"]];
const AUD_INICIO = [["subito","Súbito"],["gradual","Gradual"],["od","OD"],["oe","OE"],["bilateral","Bilateral"]];
const AUD_EVOLUCAO = [["estavel","Estável"],["piora","Piora"],["flutuante","Flutuante"],["melhora","Melhora"]];
const AUD_HIST_AUDITIVA = [["avaliacaoAnterior","Já realizou avaliação audiológica anteriormente"],["perdaAuditivaPrevia","Já apresentou perda auditiva"],["usaAparelho","Usa aparelho auditivo"],["cirurgiaOuvido","Já realizou cirurgia de ouvido"],["infeccoesOtites","História de infecções de ouvido/otites"],["perfuracaoTimpanica","Perfuração de membrana timpânica"],["exposicaoRuido","Exposição a ruído intenso"],["usaProtecaoTrabalho","Usa proteção auditiva no trabalho"]];
const AUD_SINTOMAS = [["zumbido","Zumbido"],["tontura","Tontura"],["otalgia","Otalgia (dor)"],["otorreia","Otorreia (secreção)"],["plenitude","Plenitude auricular"],["hiperacusia","Hiperacusia"],["misofonia","Misofonia"]];
const AUD_HIST_MEDICA = [["diabetes","Diabetes"],["hipertensao","Hipertensão"],["cardiovasculares","Doenças cardiovasculares"],["tireoide","Alterações da tireoide"],["neurologicas","Doenças neurológicas"],["renais","Doenças renais"],["metabolicas","Doenças metabólicas"],["avc","AVC"],["traumatismoCraniano","Traumatismo craniano"],["cirurgiaNeurologica","Cirurgia neurológica"],["covidInfeccoes","COVID-19 ou outras infecções relevantes"]];
const AUD_OCUPACIONAL = [["trabalhouExpostoRuido","Já trabalhou exposto a ruído"]];
const AUD_NIVEL_RUIDO = [["leve","Leve"],["moderado","Moderado"],["intenso","Intenso"]];
const AUD_EPI = [["nunca","Nunca"],["asVezes","Às vezes"],["sempre","Sempre"]];
const AUD_OUTRAS_EXPOSICOES = [["produtosQuimicos","Produtos químicos"],["vibracao","Vibração"],["solventes","Solventes"],["musicaAmplificada","Música amplificada"]];
const AUD_SITUACOES_DIFICULDADE = [["conversaIndividual","Conversa individual"],["conversaGrupo","Conversa em grupo"],["ambienteRuidoso","Ambiente ruidoso"],["televisao","Televisão"],["telefone","Telefone"],["reunioes","Reuniões"],["igrejaPalestras","Igreja/palestras"],["restaurantes","Restaurantes"],["direcaoRua","Direção/rua"]];
const AUD_PERCEPCAO = [["normal","Normal"],["leveDificuldade","Leve dificuldade"],["moderadaDificuldade","Moderada dificuldade"],["grandeDificuldade","Grande dificuldade"]];
const AUD_OUVIDO_PIOR = [["od","OD"],["oe","OE"],["iguais","Iguais"],["naoSabe","Não sabe"]];
const AUD_INTERFERENCIA = [["naoInterfere","Não interfere"],["pouco","Pouco"],["moderadamente","Moderadamente"],["muito","Muito"]];

const ANAMNESE_ADULTO_SECTIONS = [
  { title: "Identificação (dados audiológicos)", children: [ grp("Contato e encaminhamento", [ txt("telefone","Telefone"), txt("profissao","Profissão"), txt("encaminhamento","Encaminhamento") ]) ] },
  { title: "Queixa Principal", children: [ chk("motivo", AUD_QUEIXA_ADULTO, { path: "outro", label: "Outro" }), chk("inicio", AUD_INICIO), chk("evolucao", AUD_EVOLUCAO) ] },
  { title: "História Auditiva", children: [ chk("checks", AUD_HIST_AUDITIVA), ta("detalhes","Detalhes (quando, qual ouvido, modelo/tempo de uso de AASI, tipo de cirurgia, contexto de exposição a ruído)",3) ] },
  { title: "Sintomas Otológicos", children: [ chk("checks", AUD_SINTOMAS), ta("outros","Outros sintomas",2) ] },
  { title: "História Médica", children: [ chk("checks", AUD_HIST_MEDICA), ta("outras","Outras",2) ] },
  { title: "Uso de Medicamentos", children: [ ta("medicamentos","Utiliza medicamentos atualmente? Quais?",2) ] },
  { title: "História Familiar", children: [ txt("quem","Familiares com perda auditiva — quem?"), chk("tipoPerda",[["congenita","Congênita"],["adquirida","Adquirida"],["naoSabe","Não sabe"]]) ] },
  { title: "História Ocupacional", children: [
    grp("Exposição", [ txt("profissaoAtual","Profissão atual"), txt("tempoExposicao","Tempo de exposição") ]),
    chk("checks", AUD_OCUPACIONAL), chk("nivelRuido", AUD_NIVEL_RUIDO), chk("usoEPI", AUD_EPI), chk("outrasExposicoes", AUD_OUTRAS_EXPOSICOES),
  ]},
  { title: "Hábitos e Comunicação", children: [
    chk("usoFones",[["nao","Não usa fones"],["raramente","Raramente"],["diariamente","Diariamente"],["variasHorasDia","Várias horas/dia"]]),
    chk("situacoesDificuldade", AUD_SITUACOES_DIFICULDADE),
  ]},
  { title: "Percepção do Paciente", children: [ chk("percepcao", AUD_PERCEPCAO), chk("ouvidoPior", AUD_OUVIDO_PIOR), chk("interferencia", AUD_INTERFERENCIA) ] },
  { title: "Observações do Fonoaudiólogo", children: [ ta("obs","Observações",4) ] },
];

const AUD_QUEIXA_INFANTIL = [["dificuldadeParaOuvir","Dificuldade para ouvir"],["atrasoFalaLinguagem","Atraso de fala/linguagem"],["dificuldadeEscolar","Dificuldade escolar"],["otites","Otites"],["suspeitaPerdaAuditiva","Suspeita de perda auditiva"],["avaliacaoRotina","Avaliação de rotina"],["acompanhamentoAudiologico","Acompanhamento audiológico"]];
const AUD_GESTACAO = [["semIntercorrencias","Sem intercorrências"]];
const AUD_PARTO = [["normal","Normal"],["cesarea","Cesárea"],["forceps","Fórceps"]];
const AUD_NEONATAL = [["semIntercorrencias","Sem intercorrências"],["prematuridade","Prematuridade"],["utiNeonatal","UTI neonatal"],["oxigenoterapia","Oxigenoterapia/ventilação"],["ictericiaFototerapia","Icterícia/Fototerapia"],["medicamentosOtotoxicos","Uso de medicamentos ototóxicos"]];
const AUD_TAN = [["passou","Passou"],["falhou","Falhou"],["naoRealizou","Não realizou"],["naoSabe","Não sabe"]];
const AUD_HIST_INFANTIL = [["otitesRepeticao","Otites de repetição"],["dorOuSecrecao","Dor ou secreção no ouvido"],["cirurgiaTuboVentilacao","Cirurgia de ouvido / Tubo de ventilação"],["traumaAcustico","Trauma acústico / exposição a ruído intenso"],["usaAparelhoImplante","Uso de aparelho auditivo / implante coclear"],["historicoFamiliar","Histórico familiar de perda auditiva"]];
const AUD_DESENV_GLOBAL = [["adequado","Adequado"],["atrasado","Atrasado"],["naoSabeInformar","Não sabe informar"]];
const AUD_BALBUCIO = [["sim","Sim"],["nao","Não"],["naoSabe","Não sabe"]];
const AUD_COMPREENDE = [["sim","Sim"],["nao","Não"],["parcialmente","Parcialmente"]];
const AUD_APRESENTA = [["trocasNaFala","Trocas na fala"],["dificuldadeCompreensao","Dificuldade de compreensão"],["dificuldadeComunicacao","Dificuldade de comunicação"],["atrasoLinguagem","Atraso de linguagem"],["fazTratamentoFono","Faz tratamento fonoaudiológico"]];
const AUD_COMPORT_AUDITIVO = [["naoRespondeQuandoChamado","Não responde quando chamado"],["dificuldadeAmbienteRuidoso","Dificuldade para ouvir em ambientes ruidosos"],["solicitaRepeticao","Solicita repetição frequentemente"],["dificuldadeLocalizarSom","Dificuldade para identificar de onde vem o som"],["aumentaVolumeTV","Aumenta o volume da TV/dispositivos"],["demonstraIncomodoSons","Demonstra incômodo com sons"],["aproximaMuitoTV","Aproxima-se muito da TV"],["hipersensibilidadeSons","Hipersensibilidade a sons"],["confundePalavras","Confunde palavras"],["naoApresentaQueixas","Não apresenta queixas auditivas"]];
const AUD_HIST_MEDICA_INFANTIL = [["meningite","Meningite"],["convulsoes","Convulsões"],["traumatismoCraniano","Traumatismo craniano"],["doencasNeurogeneticas","Doenças neurológicas/genéticas"],["internacoes","Internações"]];
const AUD_EXAMES_ANTERIORES = [["nenhuma","Nenhuma"],["audiometria","Audiometria"],["imitanciometria","Imitanciometria"],["eoa","Emissões otoacústicas (EOA)"],["peateBera","PEATE / BERA"],["peaeee","PEAEEE"],["avaliacaoComportamental","Avaliação comportamental"]];

const ANAMNESE_INFANTIL_SECTIONS = [
  { title: "Identificação e Motivo da Avaliação", children: [
    grp("Contexto escolar", [ txt("escolaCreche","Escola/Creche"), txt("serieAno","Série/Ano") ]),
    chk("queixa", AUD_QUEIXA_INFANTIL, { path: "outro", label: "Outro" }), ta("obs","Observações",2),
  ]},
  { title: "Gestação, Parto e Período Neonatal", children: [
    grp("Nascimento", [ txt("idadeGestacional","Idade gestacional (semanas)"), txt("pesoNascer","Peso ao nascer (g)") ]),
    chk("gestacao", AUD_GESTACAO, { path: "intercorrenciasQuais", label: "Com intercorrências — quais?" }),
    chk("parto", AUD_PARTO),
    chk("periodoNeonatal", AUD_NEONATAL, { path: "outrosNeonatal", label: "Outros" }),
    chk("tan", AUD_TAN),
  ]},
  { title: "História Auditiva e Otológica", children: [
    chk("checks", AUD_HIST_INFANTIL),
    grp("Detalhes", [ txt("frequenciaOtites","Frequência das otites"), txt("parentescoFamiliar","Parentesco (perda auditiva familiar)") ]),
    ta("outrasInformacoes","Outras informações relevantes",2),
  ]},
  { title: "Desenvolvimento de Fala e Linguagem", children: [
    chk("desenvolvimentoGlobal", AUD_DESENV_GLOBAL),
    chk("balbucioIdadeEsperada", AUD_BALBUCIO),
    grp("Marcos", [ txt("primeirasPalavrasMeses","Primeiras palavras (meses)"), txt("frasesMeses","Frases (meses)") ]),
    chk("compreendeComandos", AUD_COMPREENDE),
    chk("apresenta", AUD_APRESENTA),
    ta("obs","Observações",2),
  ]},
  { title: "Comportamento Auditivo", children: [ chk("checks", AUD_COMPORT_AUDITIVO), ta("outrasObservacoes","Outras observações",2) ] },
  { title: "História Médica Relevante", children: [
    chk("checks", AUD_HIST_MEDICA_INFANTIL),
    chk("usoContinuoMedicamentos",[["nao","Não"],["sim","Sim"]], { path: "quais", label: "Quais" }),
    ta("obs","Observações",2),
  ]},
  { title: "Avaliações/Exames Audiológicos Anteriores", children: [
    chk("checks", AUD_EXAMES_ANTERIORES, { path: "outros", label: "Outros" }),
    grp("Resultado", [ txt("datas","Data(s)"), txt("resultados","Resultado(s)") ]),
  ]},
  { title: "Informações Adicionais do Responsável", children: [ ta("obs","Informações adicionais",3) ] },
];

const MEATO_ACHADOS = [["livre","Meato acústico externo livre"],["cerume","Presença de cerume"],["corpoEstranho","Corpo estranho"],["secrecao","Secreção"],["colabamentoCanal","Colabamento do canal"],["outrasAlteracoes","Outras alterações otoscópicas"]];
const MEATO_VIABILIDADE = [["possivelAmbas","Possível realizar audiometria em ambas as orelhas"],["possivelParcial","Possível apenas em uma orelha"],["naoPossivel","Não possível — encaminhar antes da audiometria (ex.: remoção de cerume)"]];

const AUDIOMETRIA_SECTIONS = [
  { title: "Meatoscopia (triagem inicial)", children: [
    grp("Orelha direita", [ chk("checks", MEATO_ACHADOS), txt("obs","Obs.") ]),
    grp("Orelha esquerda", [ chk("checks", MEATO_ACHADOS), txt("obs","Obs.") ]),
    chk("viabilidadeAudiometria", MEATO_VIABILIDADE),
    ta("conduta","Conduta (ex.: encaminhar para remoção de cerume antes de prosseguir)",2),
  ]},
  { title: "Equipamentos e Calibração", children: [
    grp("Audiômetro", [ txt("audiometroModelo","Modelo/Marca"), txt("audiometroCalibradoEm","Calibrado em") ]),
    grp("Imitanciômetro", [ txt("imitanciometroModelo","Modelo/Marca"), txt("imitanciometroCalibradoEm","Calibrado em") ]),
  ]},
  { title: "Audiometria Tonal Limiar (via aérea / via óssea) e Laudo", children: [
    cst("audiograma", blankAudiograma, AudiogramEditor, AudiogramReadView, audiogramTextLines),
  ]},
  { title: "Logoaudiometria (SRT/LRF · LDF · IPRF)", children: [
    cst("logoaudio", blankLogoaudio, LogoaudioEditor, LogoaudioReadView, logoaudioTextLines),
  ]},
  { title: "Imitanciometria — Timpanometria", children: [
    cst("timpanometria", blankTimpanometria, TimpanometriaEditor, TimpanometriaReadView, timpanometriaTextLines),
  ]},
  { title: "Imitanciometria — Limites dos Reflexos Acústicos", children: [
    cst("reflexos", blankReflexos, ReflexosEditor, ReflexosReadView, reflexosTextLines),
    ta("anotacoesRelevantes","Anotações relevantes",3),
  ]},
  { title: "Potenciais Evocados Auditivos (PEATE / BERA)", children: [
    grp("Parâmetros do estímulo", [
      chk("estimulo",[["clique","Clique"],["toneBurst","Tone burst"]]),
      txt("taxaApresentacao","Taxa de apresentação (estímulos/s)"),
      txt("intensidadeInicial","Intensidade inicial (dB NA)"),
      txt("numeroEstimulos","Nº de estímulos"),
    ]),
    grp("Ondas e interpicos — Orelha direita (ms)", [
      txt("odOndaI","Onda I"), txt("odOndaIII","Onda III"), txt("odOndaV","Onda V"),
      txt("odInterpicoI_III","Interpico I-III"), txt("odInterpicoIII_V","Interpico III-V"), txt("odInterpicoI_V","Interpico I-V"),
    ]),
    grp("Ondas e interpicos — Orelha esquerda (ms)", [
      txt("oeOndaI","Onda I"), txt("oeOndaIII","Onda III"), txt("oeOndaV","Onda V"),
      txt("oeInterpicoI_III","Interpico I-III"), txt("oeInterpicoIII_V","Interpico III-V"), txt("oeInterpicoI_V","Interpico I-V"),
    ]),
    grp("Limiar eletrofisiológico e classificação", [
      txt("limiarOD","Limiar eletrofisiológico OD (dB NA)"), txt("limiarOE","Limiar eletrofisiológico OE (dB NA)"),
      txt("classificacaoOD","Classificação OD"), txt("classificacaoOE","Classificação OE"),
    ]),
    ta("obs","Observações / referência",3),
  ]},
  { title: "Emissões Otoacústicas Evocadas (EOA)", children: [
    chk("tipo",[["transientes","Transientes (TEOAE)"],["produtoDistorcao","Produto de distorção (DPOAE)"]]),
    chk("resultadoOD",[["presentePassa","OD — Presente / Passa"],["ausenteFalha","OD — Ausente / Falha"]]),
    chk("resultadoOE",[["presentePassa","OE — Presente / Passa"],["ausenteFalha","OE — Ausente / Falha"]]),
    ta("obs","Relação sinal-ruído por frequência, condições do exame, observações",3),
  ]},
  { title: "Eletrococleografia (ECoG)", children: [
    grp("Parâmetros do registro", [ txt("estimulo","Estímulo"), txt("intensidade","Intensidade (dB NA)") ]),
    grp("Orelha direita", [ txt("odSP","SP — potencial de somação (µV)"), txt("odAP","AP — potencial de ação (µV)"), txt("odRazao","Razão SP/AP"), txt("odLatencia","Latência (ms)") ]),
    grp("Orelha esquerda", [ txt("oeSP","SP — potencial de somação (µV)"), txt("oeAP","AP — potencial de ação (µV)"), txt("oeRazao","Razão SP/AP"), txt("oeLatencia","Latência (ms)") ]),
    chk("indicacao",[["suspeitaMeniere","Suspeita de Doença de Ménière"],["hidropisiaEndolinfatica","Suspeita de hidropisia endolinfática"],["outraIndicacao","Outra indicação"]]),
    ta("obs","Observações / referência",2),
  ]},
  { title: "Audiometria Comportamental Infantil (BOA / VRA / COR)", children: [
    chk("tipoTeste",[["boa","BOA — Observação comportamental (< 6 meses)"],["vra","VRA — Reforço visual (6–24 meses)"],["cor","COR — Resposta de orientação condicionada"],["playAudiometria","Audiometria por brincar (Play, > 2,5 anos)"]]),
    grp("Contexto", [ txt("idadeMeses","Idade (meses)"), txt("estadoComportamental","Estado comportamental (alerta/sono leve/choro)") ]),
    chk("reforco",[["visualLuminoso","Reforço visual luminoso"],["brinquedoMecanico","Brinquedo mecânico animado"],["social","Reforço social (elogio/aplauso)"]]),
    grp("Nível mínimo de resposta — campo livre (dB NA)", [
      txt("nmr500","500 Hz"), txt("nmr1000","1000 Hz"), txt("nmr2000","2000 Hz"), txt("nmr4000","4000 Hz"),
    ]),
    grp("Nível mínimo de resposta — por orelha, se fones tolerados (dB NA)", [
      txt("odNMR","OD"), txt("oeNMR","OE"),
    ]),
    ta("obsConfiabilidade","Confiabilidade do condicionamento e observações",2),
  ]},
];

const AREA_AUDIOLOGIA_ADULTO = {
  label: "Audiologia — Adulto",
  disciplina: "Estágio em Audiologia",
  icon: Ear,
  sections: [...AUDIOMETRIA_SECTIONS],
};
const AREA_AUDIOLOGIA_INFANTIL = {
  label: "Audiologia — Infantil",
  disciplina: "Estágio em Audiologia",
  icon: Ear,
  sections: [...AUDIOMETRIA_SECTIONS],
};

/* ---------------------------------------------------------------------
   Area 3 — Fonoaudiologia em Gerontologia
   Grounded in: PARD (Padovani et al.), EAT-10, atuação fonoaudiológica
   em disfagia (CFFa/2024), presbifonia/presbiacusia na literatura geral.
------------------------------------------------------------------------ */
const GER_GERAL = [["autonomiaAVDs","Autonomia nas AVDs"],["polifarmacia","Uso de múltiplas medicações"],["quedasRecentes","Quedas recentes"],["alteracoesVisuais","Alterações visuais"],["mobilidadeReduzida","Mobilidade reduzida"]];
const GER_COGN = [["orientacaoTemporoEspacial","Orientação têmporo-espacial preservada"],["queixaMemoria","Queixa de memória"],["dificuldadeAtencao","Dificuldade de atenção"],["alteracaoLinguagem","Alteração de linguagem"]];
const GER_VOZ = [["rouquidao","Rouquidão / voz áspera"],["fadigaVocal","Fadiga vocal"],["reducaoProjecao","Redução da projeção vocal"],["tremorVocal","Tremor vocal"]];
const GER_AUD = [["queixaAuditiva","Queixa auditiva"],["usoAASI","Uso de AASI"],["dificuldadeRuido","Dificuldade em ambiente ruidoso"]];
const GER_DISF = [["tosseAoDeglutir","Tosse ao deglutir"],["engasgos","Engasgos"],["vozMolhada","Voz molhada após deglutir"],["perdaDePeso","Perda de peso recente"],["tempoAumentado","Tempo aumentado para se alimentar"],["mudancaConsistencia","Mudança espontânea de consistência alimentar"]];
const GER_RISCO = [["semRisco","Sem risco identificado"],["riscoLeve","Risco leve"],["riscoModerado","Risco moderado"],["riscoSevero","Risco severo"]];
const GER_ORAL = [["usoProtese","Uso de prótese dentária"],["proteseMalAdaptada","Prótese mal adaptada"],["perdaDentaria","Perda dentária relevante"],["dificuldadeMastigatoria","Dificuldade mastigatória"]];

/* ---------------------------------------------------------------------
   Area — Processamento Auditivo Central (PAC)
   Grounded in: Guia de Orientação Avaliação e Intervenção no PAC (CFFa),
   Guia prático para avaliação do PAC (CRFa6) — bateria comportamental
   padrão: Fala com Ruído, Dicótico de Dígitos, SSW, Padrão de Frequência/
   Duração, GIN, RGDT, MLD, SSI/PSI, Localização Sonora, Memória Sequencial.
------------------------------------------------------------------------ */
const PAC_CRITERIOS = [["audicaoNormalBilateral","Audição normal bilateral confirmada"],["qiAdequado","Capacidade cognitiva/QI compatível com a tarefa"],["semAlteracaoSignificativa","Sem alteração de linguagem/cognição que impeça o teste"]];
const PAC_HABILIDADES = [["fechamentoAuditivo","Fechamento auditivo"],["figuraFundo","Figura-fundo para sons verbais"],["atencaoSeletiva","Atenção seletiva"],["integracaoBinaural","Integração binaural"],["separacaoBinaural","Separação binaural"],["ordenacaoTemporal","Ordenação temporal"],["resolucaoTemporal","Resolução temporal"],["localizacaoSonora","Localização sonora"]];

const AREA_PAC = {
  label: "Processamento Auditivo Central (PAC)",
  disciplina: "Estágio em Processamento Auditivo Central",
  icon: Activity,
  sections: [
    { title: "Identificação e Encaminhamento", children: [
      grp("Contexto", [ txt("idadeCronologica","Idade cronológica"), txt("escolaridade","Escolaridade") ]),
      ta("motivoAvaliacao","Motivo do encaminhamento",2),
    ]},
    { title: "Critérios de Inclusão", children: [ chk("checks", PAC_CRITERIOS), ta("obs","Obs.",2) ] },
    { title: "Localização Sonora e Memória Sequencial", children: [
      grp("Localização Sonora", [ txt("resultado","Resultado") ]),
      grp("Memória Sequencial", [ txt("verbal","Sons verbais"), txt("naoVerbal","Sons não verbais") ]),
    ]},
    { title: "Testes Monóticos de Baixa Redundância", children: [
      grp("Fala com Ruído (FR)", [ txt("od","% acertos OD"), txt("oe","% acertos OE") ]),
      grp("Fala Filtrada", [ txt("od","% acertos OD"), txt("oe","% acertos OE") ]),
      grp("SSI/PSI — Mensagem Competitiva (MCI/MCC)", [ txt("mci","% acertos MCI"), txt("mcc","% acertos MCC") ]),
    ]},
    { title: "Testes Dicóticos", children: [
      grp("Dicótico de Dígitos (TDD)", [ txt("od","% acertos OD"), txt("oe","% acertos OE") ]),
      grp("SSW (Staggered Spondaic Word)", [ txt("od","% acertos OD"), txt("oe","% acertos OE") ]),
      grp("Dicótico Não-verbal", [ txt("od","% acertos OD"), txt("oe","% acertos OE") ]),
    ]},
    { title: "Testes de Padrão Temporal", children: [
      grp("Padrão de Frequência (TPF)", [ txt("od","% acertos OD"), txt("oe","% acertos OE") ]),
      grp("Padrão de Duração (TPD)", [ txt("od","% acertos OD"), txt("oe","% acertos OE") ]),
    ]},
    { title: "Resolução Temporal e Interação Binaural", children: [
      grp("GIN (Gap in Noise)", [ txt("od","Limiar OD (ms)"), txt("oe","Limiar OE (ms)") ]),
      grp("RGDT (Random Gap Detection Test)", [ txt("limiar","Limiar (ms)") ]),
      grp("MLD (Masking Level Difference)", [ txt("valor","Valor (dB)") ]),
    ]},
    { title: "Interpretação — Habilidades Auditivas Alteradas", children: [ chk("checks", PAC_HABILIDADES), ta("obs","Síntese diagnóstica e correlação com queixa/desempenho escolar",3) ] },
  ],
};


/* ---------------------------------------------------------------------
   Area — Motricidade Orofacial Infantil
   Grounded in domínios padrão de avaliação miofuncional orofacial
   (histórico de hábitos orais, estruturas em repouso, frênulos,
   respiração, mastigação, deglutição, fala e oclusão) — mesma estrutura
   usada em protocolos como MBGR e OMES-B.
------------------------------------------------------------------------ */
const MOI_HABITOS = [["succaoDedo","Sucção de dedo"],["usoChupeta","Uso de chupeta"],["mamadeiraProlongada","Uso de mamadeira prolongado"],["onicofagia","Onicofagia"],["bruxismo","Bruxismo"],["respiracaoOral","Respiração oral"],["roncos","Roncos"],["apneiaDoSono","Suspeita de apneia do sono"]];
const MOI_LABIOS = [["competentes","Competentes"],["incompetentes","Incompetentes"],["hipotonicos","Hipotônicos"],["hipertonicos","Hipertônicos"]];
const MOI_LINGUA = [["posturaAdequada","Postura adequada"],["interposta","Interposta"],["protrusa","Protrusa"]];
const MOI_BOCHECHAS = [["tonusAdequado","Tônus adequado"],["hipotonicas","Hipotônicas"],["hipertonicas","Hipertônicas"]];
const MOI_PALATO = [["normal","Normal"],["ogival","Ogival"],["fissura","Fissura"]];
const MOI_FACE = [["harmonica","Face harmônica"],["retrognatismo","Retrognatismo"],["prognatismo","Prognatismo"],["assimetriaFacial","Assimetria facial"]];
const MOI_FRENULOS = [["linguaNormal","Frênulo lingual normal"],["linguaCurto","Frênulo lingual curto / anquiloglossia"],["labialAlterado","Frênulo labial superior alterado"]];
const MOI_RESPIRACAO = [["nasal","Nasal"],["oral","Oral"],["mista","Mista"]];
const MOI_MASTIGACAO = [["bilateralAlternada","Bilateral alternada"],["unilateralPreferencial","Unilateral preferencial"],["incisal","Incisal / anteriorizada"]];
const MOI_DEGLUTICAO = [["padraoAtipico","Padrão adaptado / atípico"],["projecaoLingual","Projeção lingual"],["participacaoPerioral","Participação da musculatura perioral"],["ruidos","Ruídos à deglutição"]];
const MOI_FALA = [["compativelTipico","Articulação compatível com padrão típico"],["distorcoesEstruturais","Distorções compatíveis com alteração estrutural"],["ceceio","Ceceio"]];
const MOI_OCLUSAO = [["mordidaAbertaAnterior","Mordida aberta anterior"],["mordidaCruzada","Mordida cruzada"],["sobressaliencia","Sobressaliência aumentada"],["classeI","Classe I de Angle (referida)"],["classeII","Classe II de Angle (referida)"],["classeIII","Classe III de Angle (referida)"]];

const AREA_MOTRICIDADE_INFANTIL = {
  label: "Motricidade Orofacial Infantil",
  disciplina: "Estágio em Motricidade Orofacial (Infantil)",
  icon: Smile,
  sections: [
    { title: "Estruturas Orofaciais em Repouso", children: [
      grp("Lábios", [ chk("checks", MOI_LABIOS) ]),
      grp("Língua", [ chk("checks", MOI_LINGUA) ]),
      grp("Bochechas", [ chk("checks", MOI_BOCHECHAS) ]),
      grp("Palato", [ chk("checks", MOI_PALATO) ]),
      grp("Face / Mandíbula", [ chk("checks", MOI_FACE) ]),
    ]},
    { title: "Frênulos", children: [ chk("checks", MOI_FRENULOS), ta("obs","Obs.",2) ] },
    { title: "Respiração", children: [ chk("checks", MOI_RESPIRACAO), ta("obs","Padrão e tipo respiratório",2) ] },
    { title: "Mastigação", children: [ chk("checks", MOI_MASTIGACAO), ta("obs","Obs.",2) ] },
    { title: "Deglutição", children: [ chk("checks", MOI_DEGLUTICAO), ta("obs","Obs.",2) ] },
    { title: "Fala / Articulação Relacionada à Motricidade Orofacial", children: [ chk("checks", MOI_FALA), ta("obs","Obs.",2) ] },
    { title: "Oclusão Dentária (achados referidos/observados)", children: [ chk("checks", MOI_OCLUSAO), ta("obs","Obs.",2) ] },
    { title: "Diagnóstico Funcional e Conduta", children: [ ta("obs","Síntese diagnóstica e plano terapêutico de motricidade orofacial",4) ] },
  ],
};


const AREA_GERONTOLOGIA = {
  label: "Fonoaudiologia em Gerontologia",
  disciplina: "Estágio em Gerontologia",
  icon: HeartPulse,
  sections: [
    { title: "Aspectos Gerais do Envelhecimento", children: [ chk("checks", GER_GERAL), ta("obs", "Obs.", 2) ] },
    { title: "Rastreio Cognitivo e Comunicação", children: [ chk("checks", GER_COGN), ta("obs", "Observações (ex.: resultado de rastreio cognitivo aplicado)", 3) ] },
    { title: "Voz (Presbifonia)", children: [ chk("checks", GER_VOZ), ta("obs", "Obs.", 2) ] },
    { title: "Audição", children: [ chk("checks", GER_AUD), ta("obs", "Obs.", 2) ] },
    { title: "Deglutição — Triagem de Risco para Disfagia", children: [
      chk("checks", GER_DISF),
      txt("eat10", "Escore EAT-10 (se aplicado)"),
      chk("classificacao", GER_RISCO),
      ta("obs", "Obs.", 3),
    ]},
    { title: "Saúde Oral e Mastigação", children: [ chk("checks", GER_ORAL), ta("obs", "Obs.", 2) ] },
    { title: "Impacto na Qualidade de Vida e Participação Social", children: [ ta("obs", "Observações", 3) ] },
  ],
};

/* ---------------------------------------------------------------------
   Area 4 — Voz
------------------------------------------------------------------------ */
const VOZ_QUEIXAS = [["rouquidao","Rouquidão"],["fadigaVocal","Fadiga vocal"],["perdaDeVoz","Perda de voz / afonia"],["dorAoFalar","Dor ao falar"],["pigarro","Pigarro"],["tosseSeca","Tosse seca"],["usoProfissional","Uso profissional da voz"]];
const VOZ_RESP = [["toracico","Tipo respiratório torácico"],["abdominal","Tipo respiratório abdominal"],["misto","Tipo respiratório misto"],["coordenacaoAlterada","Coordenação pneumofonoarticulatória alterada"],["tensaoCervical","Tensão cervical"]];
const VOZ_RESSON = [["equilibrada","Ressonância equilibrada"],["hipernasal","Hipernasalidade"],["hiponasal","Hiponasalidade"],["articulacaoImprecisa","Articulação imprecisa"]];
const VOZ_ATAQUE = [["brusco","Ataque vocal brusco"],["soproso","Ataque vocal soproso"],["pitchInadequado","Pitch inadequado"],["loudnessInadequada","Loudness inadequada"]];

const AREA_VOZ = {
  label: "Voz",
  disciplina: "Estágio em Voz",
  icon: Mic2,
  sections: [
    { title: "Anamnese Vocal", children: [
      chk("checks", VOZ_QUEIXAS), txt("tempoQueixa", "Tempo da queixa"),
      ta("habitos", "Hábitos vocais (hidratação, tabagismo, etilismo, cafeína, gritar)", 3),
    ]},
    { title: "Avaliação Perceptivo-Auditiva (escala GRBASI, 0–3)", children: [
      grp("Parâmetros", [ txt("g","Grau geral (G)"), txt("r","Rugosidade (R)"), txt("b","Soprosidade (B)"), txt("a","Astenia (A)"), txt("s","Tensão (S)"), txt("i","Instabilidade (I)") ]),
    ]},
    { title: "Aspectos Respiratórios e Posturais", children: [ chk("checks", VOZ_RESP), ta("obs", "Obs.", 2) ] },
    { title: "Ressonância e Articulação", children: [ chk("checks", VOZ_RESSON), ta("obs", "Obs.", 2) ] },
    { title: "Ataque Vocal, Pitch e Loudness", children: [ chk("checks", VOZ_ATAQUE), ta("obs", "Obs.", 2) ] },
    { title: "Exame Laríngeo (referido)", children: [ ta("achados", "Achados de laringoscopia / videoestroboscopia, se disponíveis", 3) ] },
    { title: "Impacto Vocal", children: [ ta("obs", "Impacto nas atividades diárias e profissionais", 3) ] },
  ],
};

/* ---------------------------------------------------------------------
   Area 5 — Disfagia / Motricidade Orofacial no Adulto
   Grounded in: PARD, EAT-10, Resolução CFFa 719/2023, POP UR-FONO-018.
------------------------------------------------------------------------ */
const DISF_TRIAGEM = [["tosseAoEngolir","Tosse ao engolir"],["engasgo","Engasgo"],["vozMolhada","Voz molhada após deglutir"],["dificuldadeMastigar","Dificuldade de mastigar"],["sialorreia","Sialorreia"],["perdaDePeso","Perda de peso"]];
const DISF_CLINICO = [["nivelConscienciaAdequado","Nível de consciência adequado"],["viaAereaProtegida","Via aérea protegida"],["traqueostomia","Traqueostomia"],["ventilacaoMecanica","Ventilação mecânica"],["ausculaCervicalAlterada","Ausculta cervical alterada"]];
const DISF_TESTES = [["aguaAlterado","Teste da água alterado"],["pastosoAlterado","Teste de pastoso alterado"],["solidoAlterado","Teste de sólido alterado"]];
const DISF_SINAIS = [["tosseAposDeglutir","Tosse após deglutir"],["engasgoAposDeglutir","Engasgo após deglutir"],["alteracaoVocalAposDeglutir","Alteração vocal após deglutir"],["saturacaoReduzida","Redução da saturação de O2"]];
const DISF_VIA = [["viaOralExclusiva","Via oral exclusiva"],["viaOralComRestricoes","Via oral com restrições"],["viaAlternativa","Via alternativa (sonda / gastrostomia)"]];
const DISF_MO = [["labios","Lábios"],["lingua","Língua"],["bochechas","Bochechas"],["palato","Palato"],["mandibula","Mandíbula"]];

const AREA_DISFAGIA = {
  label: "Disfagia / Motricidade Orofacial no Adulto",
  disciplina: "Estágio em Disfagia / Motricidade Orofacial (Adulto)",
  icon: Utensils,
  sections: [
    { title: "Triagem de Risco (sinais clínicos / EAT-10)", children: [ chk("checks", DISF_TRIAGEM), txt("eat10", "Escore EAT-10"), ta("obs", "Obs.", 2) ] },
    { title: "Estado Clínico e Respiratório", children: [ chk("checks", DISF_CLINICO), ta("obs", "Obs.", 2) ] },
    { title: "Avaliação Clínica da Deglutição", children: [ chk("checks", DISF_TESTES), ta("obs", "Obs.", 2) ] },
    { title: "Sinais de Penetração / Aspiração", children: [ chk("checks", DISF_SINAIS), ta("obs", "Obs.", 2) ] },
    { title: "Via de Alimentação", children: [ chk("checks", DISF_VIA), txt("consistenciaLiberada", "Consistência liberada") ] },
    { title: "Motricidade Orofacial (estruturas)", children: [ chk("checks", DISF_MO), ta("obs", "Obs.", 2) ] },
    { title: "Condutas e Encaminhamentos", children: [ ta("obs", "Condutas e encaminhamentos", 3) ] },
  ],
};

/* ---------------------------------------------------------------------
   Area 6 — Linguagem no Adulto / Neurologia (Afasiologia)
------------------------------------------------------------------------ */
const NEU_ETIO = [["avc","AVC"],["tce","TCE"],["neurodegenerativa","Doença neurodegenerativa"],["tumor","Tumor"],["outra","Outra etiologia"]];
const NEU_EXPR = [["fluenciaReduzida","Fluência reduzida"],["anomia","Anomia"],["parafasias","Parafasias"],["agramatismo","Agramatismo"],["apraxiaDeFala","Apraxia de fala"],["disartria","Disartria"]];
const NEU_COMP = [["ordensSimples","Compreensão de ordens simples"],["ordensComplexas","Compreensão de ordens complexas"],["compreensaoTextual","Compreensão textual"]];
const NEU_LE = [["leituraPreservada","Leitura preservada"],["alexia","Alexia"],["escritaPreservada","Escrita preservada"],["agrafia","Agrafia"]];
const NEU_CLASSIF = [["broca","Broca"],["wernicke","Wernicke"],["conducao","Condução"],["global","Global"],["anomica","Anômica"],["transMotora","Transcortical motora"],["transSensorial","Transcortical sensorial"]];
const NEU_COGN = [["atencao","Atenção"],["memoria","Memória"],["funcoesExecutivas","Funções executivas"],["orientacao","Orientação"]];
const NEU_CAA = [["indicacaoCAA","Indicação de CAA"],["pranchaComunicacao","Uso de prancha de comunicação"],["aplicativo","Uso de aplicativo de CAA"]];

const AREA_NEURO = {
  label: "Linguagem no Adulto / Neurologia (Afasiologia)",
  disciplina: "Estágio em Linguagem do Adulto / Neurologia",
  icon: Brain,
  sections: [
    { title: "Etiologia e Histórico Neurológico", children: [ chk("checks", NEU_ETIO), txt("tempoEvolucao", "Tempo desde o evento"), ta("historico", "Histórico", 2) ] },
    { title: "Linguagem Expressiva", children: [ chk("checks", NEU_EXPR), ta("obs", "Obs.", 2) ] },
    { title: "Linguagem Compreensiva", children: [ chk("checks", NEU_COMP), ta("obs", "Obs.", 2) ] },
    { title: "Leitura e Escrita", children: [ chk("checks", NEU_LE), ta("obs", "Obs.", 2) ] },
    { title: "Classificação da Afasia (se aplicável)", children: [ chk("checks", NEU_CLASSIF), ta("obs", "Obs.", 2) ] },
    { title: "Cognição Comunicativa", children: [ chk("checks", NEU_COGN), ta("obs", "Obs.", 2) ] },
    { title: "Comunicação Suplementar/Alternativa (CAA)", children: [ chk("checks", NEU_CAA), ta("obs", "Obs.", 2) ] },
  ],
};

/* ---------------------------------------------------------------------
   Area 7 — Neonatal (UTI Neonatal / Berçário)
   Grounded in: Lei 12.303/2010 e diretrizes de TAN (MS/CGSPD), IRDA (JCIH),
   rotina fonoaudiológica em UTIN (POP.FONO.007 — INTS).
------------------------------------------------------------------------ */
const NEO_IRDA = [["prematuridade","Prematuridade"],["baixoPeso","Baixo peso ao nascer"],["permanenciaUTIN","Permanência em UTIN > 5 dias"],["ototoxicos","Uso de medicações ototóxicas"],["historiaFamiliar","História familiar de perda auditiva"],["malformacaoCraniofacial","Malformação craniofacial"],["infeccaoCongenita","Infecção congênita"],["hiperbilirrubinemia","Hiperbilirrubinemia"]];
const NEO_REFLEXOS = [["procura","Procura"],["sucção","Sucção"],["deglutição","Deglutição"],["vomito","Vômito"],["mordida","Mordida"]];
const NEO_ESTRUTURAS = [["labios","Lábios"],["lingua","Língua"],["bochechas","Bochechas"],["mandibula","Mandíbula"],["palatoDuro","Palato duro"],["palatoMole","Palato mole"]];
const NEO_SDR = [["coordenacaoAdequada","Coordenação sucção-deglutição-respiração adequada"],["incoordenacaoSDR","Incoordenação S/D/R"],["engasgosDuranteMamada","Engasgos durante a mamada"],["quedaDeSaturacao","Queda de saturação durante alimentação"],["fadigaDuranteMamada","Fadiga durante a mamada"]];
const NEO_VIA = [["seioMaterno","Seio materno"],["mamadeira","Mamadeira"],["copo","Copo"],["sondaOrogastrica","Sonda orogástrica"],["sondaNasogastrica","Sonda nasogástrica"],["gastrostomia","Gastrostomia"]];
const NEO_TAN = [["eoaRealizada","EOA realizada"],["peateRealizado","PEATE realizado"],["resultadoPassa","Resultado: passa"],["resultadoFalha","Resultado: falha / reteste indicado"]];

const AREA_NEONATAL = {
  label: "Neonatal (UTI Neonatal / Berçário)",
  disciplina: "Estágio em Neonatologia",
  icon: Baby,
  sections: [
    { title: "Reflexos Orais", children: [ chk("checks", NEO_REFLEXOS), ta("obs","Presença, ausência ou inconsistência dos reflexos",2) ] },
    { title: "Avaliação das Estruturas Orofaciais", children: [ chk("checks", NEO_ESTRUTURAS), ta("obs","Morfologia, postura e tonicidade",2) ] },
    { title: "Coordenação Sucção-Deglutição-Respiração", children: [ chk("checks", NEO_SDR), ta("obs","Obs.",2) ] },
    { title: "Via de Alimentação", children: [ chk("checks", NEO_VIA), ta("obs","Obs.",2) ] },
    { title: "Triagem Auditiva Neonatal (TAN)", children: [ chk("checks", NEO_TAN), ta("obs","Obs.",2) ] },
    { title: "Estimulação Sensório-Motora Oral", children: [ ta("obs","Condutas de estimulação realizadas",3) ] },
  ],
};

/* ---------------------------------------------------------------------
   Area 8 — Fonoaudiologia Hospitalar (enfermaria / beira-leito)
   Grounded in: PR_162 HC-UFMG (avaliação fonoaudiológica adulto internado),
   parecer SBFa sobre atuação em UTI, FOIS (Functional Oral Intake Scale).
------------------------------------------------------------------------ */
const HOSP_QUADRO = [["ventilacaoMecanica","Ventilação mecânica"],["traqueostomia","Traqueostomia"],["viaAereaArtificial","Via aérea artificial"],["nivelConscienciaAdequado","Nível de consciência adequado"],["instabilidadeHemodinamica","Instabilidade hemodinâmica"],["contraindicacaoAvaliacao","Contraindicação clínica para avaliação"]];
const HOSP_DEGLUT = [["testeAguaAlterado","Teste da água alterado"],["testeAlimentoAlterado","Teste com alimento alterado"],["tosseReflexaPresente","Tosse reflexa presente"],["vozMolhada","Voz molhada após deglutir"]];
const HOSP_VIA = [["viaOralLiberada","Via oral liberada"],["viaOralComRestricoes","Via oral com restrições"],["sondaNasoenteral","Sonda nasoenteral"],["gastrostomia","Gastrostomia"],["nutricaoParenteral","Nutrição parenteral"]];
const HOSP_COMUNIC = [["comunicacaoOralPreservada","Comunicação oral preservada"],["usoDeGestos","Uso de gestos"],["necessidadeDeCAA","Necessidade de CAA"],["disartria","Disartria"],["afasia","Afasia"]];

const AREA_HOSPITALAR = {
  label: "Fonoaudiologia Hospitalar",
  disciplina: "Estágio em Fonoaudiologia Hospitalar",
  icon: Building2,
  sections: [
    { title: "Dados da Internação", children: [
      grp("Identificação clínica", [ txt("setorLeito","Setor / leito"), txt("diagnosticoMedico","Diagnóstico médico"), txt("dataInternacao","Data de internação"), txt("medicoSolicitante","Médico solicitante") ]),
      ta("motivoInterconsulta","Motivo da interconsulta / demanda",2),
    ]},
    { title: "Quadro Clínico Atual", children: [ chk("checks", HOSP_QUADRO), ta("obs","Obs.",2) ] },
    { title: "Avaliação da Deglutição à Beira-Leito", children: [ chk("checks", HOSP_DEGLUT), ta("obs","Obs.",2) ] },
    { title: "Via de Alimentação e Nutrição", children: [ chk("checks", HOSP_VIA), txt("consistenciaLiberada","Consistência liberada") ] },
    { title: "Comunicação", children: [ chk("checks", HOSP_COMUNIC), ta("obs","Obs.",2) ] },
    { title: "Escala Funcional de Ingesta Oral (FOIS)", children: [ txt("escoreFOIS","Escore FOIS (1–7)"), ta("justificativa","Justificativa do escore",2) ] },
    { title: "Plano de Alta e Encaminhamentos", children: [ ta("obs","Plano de alta e encaminhamentos",3) ] },
  ],
};

/* ---------------------------------------------------------------------
   Area 9 — Fonoaudiologia Escolar / Educacional
   Grounded in: literatura sobre consciência fonológica e alfabetização
   (CONFIAS, PCFO, PHF) e roteiros de avaliação fonoaudiológica escolar.
------------------------------------------------------------------------ */
const ESC_QUEIXA = [["queixaPedagogica","Queixa pedagógica"],["dificuldadeDeAprendizagem","Dificuldade de aprendizagem"],["dificuldadeDeAtencao","Dificuldade de atenção"],["repetenciaEscolar","Repetência escolar"]];
const ESC_SALA = [["participacaoAtiva","Participação ativa em sala"],["dificuldadeDeInteracao","Dificuldade de interação com colegas"],["dificuldadeComandos","Dificuldade em compreender comandos"],["dificuldadeAtencaoSustentada","Dificuldade de atenção sustentada"]];
const ESC_LINGUAGEM = [["vocabularioAdequado","Vocabulário adequado para a idade"],["estruturacaoDeFrasesAdequada","Estruturação de frases adequada"],["dificuldadeNarrativa","Dificuldade em narrativas"],["trocasNaFala","Trocas / alterações na fala"]];
const ESC_CF = [["segmentacaoSilabica","Segmentação silábica"],["rima","Rima"],["aliteracao","Aliteração"],["sinteseFonemica","Síntese fonêmica"],["analiseFonemica","Análise fonêmica"]];
const ESC_LEITURA = [["decodificacao","Decodificação"],["fluenciaLeitora","Fluência leitora"],["compreensaoLeitora","Compreensão leitora"]];
const ESC_ESCRITA = [["hipoteseAlfabetica","Hipótese alfabética estabelecida"],["trocasOrtograficas","Trocas ortográficas"],["dificuldadeSegmentacao","Dificuldade de segmentação na escrita"],["disortografia","Indícios de disortografia"]];

const AREA_ESCOLAR = {
  label: "Fonoaudiologia Escolar / Educacional",
  disciplina: "Estágio em Fonoaudiologia Escolar",
  icon: GraduationCap,
  sections: [
    { title: "Observação em Sala de Aula", children: [ chk("checks", ESC_SALA), ta("obs","Obs.",2) ] },
    { title: "Linguagem Oral", children: [ chk("checks", ESC_LINGUAGEM), ta("obs","Obs.",2) ] },
    { title: "Consciência Fonológica", children: [ chk("checks", ESC_CF), ta("obs","Habilidades desenvolvidas / em desenvolvimento",2) ] },
    { title: "Leitura", children: [ chk("checks", ESC_LEITURA), ta("obs","Obs.",2) ] },
    { title: "Escrita", children: [ chk("checks", ESC_ESCRITA), ta("obs","Obs.",2) ] },
    { title: "Orientações à Escola e Professores", children: [ ta("obs","Orientações",3) ] },
  ],
};

/* ---------------------------------------------------------------------
   Anamnese Geral — preenchida uma única vez por paciente (não se repete
   a cada atendimento). Reúne identificação, queixa, desenvolvimento,
   saúde, hábitos e contexto — o que for relevante conforme as áreas que
   acompanham o paciente.
------------------------------------------------------------------------ */
const ANAM_HIST_MEDICA_GERAL = [...AUD_HIST_MEDICA, ...AUD_HIST_MEDICA_INFANTIL];
const ANAM_APRESENTA = [["trocasNaFala","Trocas na fala"],["dificuldadeCompreensao","Dificuldade de compreensão"],["dificuldadeComunicacao","Dificuldade de comunicação"],["atrasoLinguagem","Atraso de linguagem"],["fazTratamentoFono","Já fez/faz tratamento fonoaudiológico"]];

const ANAMNESE_GERAL_SECTIONS = [
  { title: "Identificação Complementar", children: [
    grp("Contato e encaminhamento", [ txt("escolaridadeOuProfissao","Escolaridade / profissão"), txt("encaminhamento","Encaminhamento (quem indicou)"), txt("telefoneContato","Telefone de contato") ]),
  ]},
  { title: "Queixa Principal e Expectativas", children: [
    ta("queixaPrincipal","Queixa principal / motivo da avaliação",3),
    ta("expectativas","O que a família/paciente espera do acompanhamento",2),
  ]},
  { title: "História Gestacional, Parto e Período Neonatal", children: [
    grp("Nascimento", [ txt("idadeGestacional","Idade gestacional"), txt("pesoNascimento","Peso ao nascimento") ]),
    chk("gestacao", AUD_GESTACAO, { path: "intercorrenciasQuais", label: "Com intercorrências — quais?" }),
    chk("parto", AUD_PARTO),
    chk("periodoNeonatal", AUD_NEONATAL, { path: "outrosNeonatal", label: "Outros" }),
    chk("tan", AUD_TAN),
  ]},
  { title: "Desenvolvimento", children: [
    chk("desenvolvimentoGlobal", AUD_DESENV_GLOBAL),
    chk("balbucioIdadeEsperada", AUD_BALBUCIO),
    grp("Marcos de linguagem", [ txt("primeirasPalavrasMeses","Primeiras palavras (meses)"), txt("frasesMeses","Frases (meses)") ]),
    ta("desenvolvimentoMotor","Desenvolvimento motor (marcos)",2),
    chk("apresenta", ANAM_APRESENTA),
  ]},
  { title: "História de Saúde", children: [
    chk("checks", ANAM_HIST_MEDICA_GERAL),
    ta("medicamentosAtuais","Medicamentos em uso atualmente",2),
    ta("cirurgiasRealizadas","Cirurgias realizadas",2),
  ]},
  { title: "História Auditiva e Otológica", children: [
    chk("checks", AUD_HIST_AUDITIVA),
    chk("sintomas", AUD_SINTOMAS),
    ta("obs","Obs.",2),
  ]},
  { title: "Hábitos Orais", children: [ chk("checks", MOI_HABITOS), txt("tempoDuracao","Tempo de duração dos hábitos"), ta("obs","Obs.",2) ] },
  { title: "Contexto Escolar", children: [
    grp("Dados escolares", [ txt("escola","Escola"), txt("serieAno","Série / ano"), txt("professor","Professor(a)") ]),
    chk("checks", ESC_QUEIXA), ta("obs","Obs.",2),
  ]},
  { title: "História Familiar", children: [
    txt("familiaresComPerdaAuditiva","Familiares com perda auditiva — quem?"),
    chk("tipoPerdaFamiliar",[["congenita","Congênita"],["adquirida","Adquirida"],["naoSabe","Não sabe"]]),
    ta("outrasCondicoesFamiliares","Outras condições de saúde na família",2),
  ]},
  { title: "Observações Gerais", children: [ ta("obs","Observações gerais sobre o paciente",4) ] },
];
const blankAnamneseGeral = () => blankFromSections(ANAMNESE_GERAL_SECTIONS);

const AREAS = {
  linguagem_fala_mo: AREA_LINGUAGEM_FALA_MO,
  neonatal: AREA_NEONATAL,
  audiologia_adulto: AREA_AUDIOLOGIA_ADULTO,
  audiologia_infantil: AREA_AUDIOLOGIA_INFANTIL,
  pac: AREA_PAC,
  gerontologia: AREA_GERONTOLOGIA,
  voz: AREA_VOZ,
  disfagia: AREA_DISFAGIA,
  hospitalar: AREA_HOSPITALAR,
  neuro: AREA_NEURO,
  escolar: AREA_ESCOLAR,
};
const AREA_ORDER = ["linguagem_fala_mo", "neonatal", "audiologia_adulto", "audiologia_infantil", "pac", "gerontologia", "voz", "disfagia", "hospitalar", "neuro", "escolar"];
const SHORT_LABELS = {
  linguagem_fala_mo: "Linguagem/Fala/MO", neonatal: "Neonatal",
  audiologia_adulto: "Audiologia (adulto)", audiologia_infantil: "Audiologia (infantil)",
  pac: "PAC", gerontologia: "Gerontologia", voz: "Voz", disfagia: "Disfagia", hospitalar: "Hospitalar",
  neuro: "Neuro/Afasiologia", escolar: "Escolar",
};

/* ---------------------------------------------------------------------
   Record shape (area-agnostic wrapper around clinico data)
------------------------------------------------------------------------ */
function blankRecord(areaId) {
  return {
    id: null, area: areaId, criadoEm: null, atualizadoEm: null,
    identificacao: {
      instituicao: "Centro Universitário Nobre – UNIFAN",
      curso: "Fonoaudiologia",
      disciplina: AREAS[areaId].disciplina,
      estagiario: "", matricula: "", supervisor: "", crfa: "",
      data: "", horario: "", sessaoNum: "",
    },
    paciente: { nome: "", nascimento: "", idade: "", responsavel: "", hipotese: "", cid: "", tempoAcompanhamento: "" },
    clinico: blankClinico(areaId),
    procedimentos: "", evolucao: "", orientacoes: "", plano: "",
    raciocinio: { alteracao: "", evidencias: "", objetivoAlcancado: "", conduta: "" },
    estudoDeCasoSessao: { discutido: false, resumo: "" },
    validacaoSupervisao: { supervisorNome: "", crfaSupervisor: "", status: "pendente", dataValidacao: "", comentario: "" },
  };
}

const genId = () => "at_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
const slugify = (s) =>
  (s || "").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "") || "paciente";

/* ---------------------------------------------------------------------
   Small presentational primitives
------------------------------------------------------------------------ */
function Field({ label, children, span }) {
  return (
    <label className={"flex flex-col gap-1 " + (span ? "sm:col-span-" + span : "")} style={{ fontFamily: "'IBM Plex Sans', sans-serif" }}>
      <span style={{ fontSize: 12.5, color: T.inkSoft, fontWeight: 500 }}>{label}</span>
      {children}
    </label>
  );
}
const inputBase = {
  fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 14.5, color: T.ink,
  background: T.surface, border: "1px solid " + T.line, borderRadius: 6,
  padding: "8px 10px", outline: "none", width: "100%",
};
function TextInput({ value, onChange, disabled, type = "text", placeholder }) {
  return (
    <input type={type} value={value || ""} placeholder={placeholder} disabled={disabled}
      onChange={(e) => onChange(e.target.value)}
      style={{ ...inputBase, opacity: disabled ? 0.75 : 1, background: disabled ? T.surfaceSoft : T.surface }}
      onFocus={(e) => (e.target.style.borderColor = T.accent)} onBlur={(e) => (e.target.style.borderColor = T.line)} />
  );
}
function TextArea({ value, onChange, disabled, rows = 3, placeholder }) {
  return (
    <textarea value={value || ""} placeholder={placeholder} disabled={disabled} rows={rows}
      onChange={(e) => onChange(e.target.value)}
      style={{ ...inputBase, opacity: disabled ? 0.75 : 1, background: disabled ? T.surfaceSoft : T.surface, resize: "vertical", lineHeight: 1.5 }}
      onFocus={(e) => (e.target.style.borderColor = T.accent)} onBlur={(e) => (e.target.style.borderColor = T.line)} />
  );
}
function CheckRow({ options, value, onChange, disabled }) {
  return (
    <div className="flex flex-wrap gap-x-5 gap-y-2">
      {options.map(([key, label]) => (
        <label key={key} className="flex items-center gap-2 cursor-pointer select-none"
          style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 14, color: T.ink, opacity: disabled && !value?.[key] ? 0.5 : 1 }}>
          <input type="checkbox" checked={!!value?.[key]} disabled={disabled}
            onChange={() => onChange(key, !value?.[key])}
            style={{ accentColor: T.accent, width: 16, height: 16 }} />
          {label}
        </label>
      ))}
    </div>
  );
}
function SubBlock({ title, children }) {
  return (
    <div className="flex flex-col gap-2 p-3" style={{ background: T.surfaceSoft, borderRadius: 8, border: "1px solid " + T.lineSoft }}>
      <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 12.5, fontWeight: 600, color: T.accentDeep }}>{title}</div>
      {children}
    </div>
  );
}
function Section({ number, title, open, onToggle, children }) {
  return (
    <div style={{ border: "1px solid " + T.line, borderRadius: 10, background: T.surface, overflow: "hidden" }}>
      <button onClick={onToggle} className="w-full flex items-center gap-3 px-4 py-3 text-left"
        style={{ background: open ? T.accentSoft : T.surface, cursor: "pointer" }}>
        <span style={{ fontFamily: "'Fraunces', serif", fontSize: 14, fontWeight: 600, color: T.accentDeep, minWidth: 20 }}>{number}</span>
        <span style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontWeight: 600, fontSize: 14.5, color: T.ink, flex: 1 }}>{title}</span>
        {open ? <ChevronDown size={17} color={T.inkSoft} /> : <ChevronRight size={17} color={T.inkSoft} />}
      </button>
      {open && <div className="px-4 py-4 flex flex-col gap-4">{children}</div>}
    </div>
  );
}

/* ---------------------------------------------------------------------
   Audiometry custom components — table entry + auto-plotted audiogram,
   mirroring the official UNIFAN "Folha de Registro de Audiometria".
------------------------------------------------------------------------ */
const cellInput = {
  width: "100%", fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 12.5, color: T.ink,
  background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 4,
  padding: "3px 4px", outline: "none", textAlign: "center",
};
function Cell({ value, onChange, disabled }) {
  return <input value={value || ""} disabled={disabled} placeholder="—" onChange={(e) => onChange(e.target.value)} style={{ ...cellInput, opacity: disabled ? 0.6 : 1 }} />;
}
const OD_COLOR = "#C0392B";
const OE_COLOR = "#2455A4";
/* Grading scale: WHO World Report on Hearing (2021), quadritonal average
   of 500/1000/2000/4000 Hz — the current international standard. */
const SEVERITY_BANDS = [
  { from: -10, to: 20, color: "#2C6E62", label: "Normal" },
  { from: 20, to: 35, color: "#8FA83E", label: "Leve" },
  { from: 35, to: 50, color: "#D8A93B", label: "Moderada" },
  { from: 50, to: 65, color: "#D97B3B", label: "Mod. severa" },
  { from: 65, to: 80, color: "#C0533A", label: "Severa" },
  { from: 80, to: 95, color: "#8E2E2E", label: "Profunda" },
  { from: 95, to: 130, color: "#5C1F1F", label: "Completa/Total" },
];
function classifyPTA(db) {
  if (db == null || isNaN(db)) return "";
  if (db < 20) return "Normal";
  if (db < 35) return "Leve";
  if (db < 50) return "Moderada";
  if (db < 65) return "Moderadamente severa";
  if (db < 80) return "Severa";
  if (db < 95) return "Profunda";
  return "Completa/Total";
}
function computePTA(va) {
  const vals = [500, 1000, 2000, 4000].map((f) => va[f]).filter((v) => v !== "" && v != null && !isNaN(v)).map(Number);
  if (vals.length < 3) return null;
  return Math.round(vals.reduce((a, b) => a + b, 0) / vals.length);
}

/* Air-bone gap and hearing-loss-type classification (per ear), derived
   from comparing the air- and bone-conduction pure-tone averages. */
function frequencyGaps(ear) {
  return [500, 1000, 2000, 4000].map((f) => {
    const ac = parseFloat(ear.va[f]), bo = parseFloat(ear.vo[f]);
    if (isNaN(ac) || isNaN(bo)) return null;
    return { freq: f, gap: Math.round((ac - bo) * 10) / 10 };
  }).filter(Boolean);
}
function classifyLossType(ptaAC, ptaBC) {
  if (ptaAC == null) return null;
  if (ptaBC == null) return "Via óssea incompleta";
  const gap = ptaAC - ptaBC;
  if (ptaAC < 20 && ptaBC < 20) return "Normal";
  if (gap >= 10 && ptaBC < 20) return "Condutiva";
  if (gap >= 10 && ptaBC >= 20) return "Mista";
  return "Neurossensorial";
}
function LossTypeSummary({ od, oe }) {
  const ptaACOD = computePTA(od.va), ptaBCOD = computePTA(od.vo);
  const ptaACOE = computePTA(oe.va), ptaBCOE = computePTA(oe.vo);
  const gapsOD = frequencyGaps(od), gapsOE = frequencyGaps(oe);
  const avgGap = (gaps) => gaps.length ? Math.round(gaps.reduce((a, g) => a + g.gap, 0) / gaps.length) : null;
  const Row = ({ label, color, ptaAC, ptaBC, gaps }) => {
    const type = classifyLossType(ptaAC, ptaBC);
    const gap = avgGap(gaps);
    return (
      <div className="flex flex-col gap-1 p-2.5" style={{ background: T.surfaceSoft, borderRadius: 7, border: "1px solid " + T.lineSoft, flex: 1, minWidth: 220 }}>
        <div style={{ fontSize: 12, fontWeight: 600, color }}>{label}</div>
        <div style={{ fontSize: 11.5, color: T.inkSoft }}>PTA via aérea: {ptaAC != null ? `${ptaAC} dB` : "—"} · PTA via óssea: {ptaBC != null ? `${ptaBC} dB` : "—"}</div>
        <div style={{ fontSize: 11.5, color: T.inkSoft }}>Gap aéreo-ósseo médio: {gap != null ? `${gap} dB` : "—"}
          {gaps.length > 0 && <span style={{ color: T.inkFaint }}> ({gaps.map((g) => `${g.freq}Hz:${g.gap}dB`).join(" · ")})</span>}
        </div>
        {type && <div style={{ fontSize: 12, fontWeight: 600, color: type === "Normal" ? T.accent : T.warn }}>Tipo de perda (estimado): {type}</div>}
      </div>
    );
  };
  if (computePTA(od.va) == null && computePTA(oe.va) == null) return null;
  return (
    <div className="flex flex-wrap gap-3">
      <Row label="OD" color={OD_COLOR} ptaAC={ptaACOD} ptaBC={ptaBCOD} gaps={gapsOD} />
      <Row label="OE" color={OE_COLOR} ptaAC={ptaACOE} ptaBC={ptaBCOE} gaps={gapsOE} />
    </div>
  );
}

/* Auto-drafted laudo proposal, built from the computed classifications.
   Always a starting point for the student to review/edit — never final. */
function generateLaudoText(audiograma, logoaudio, timpanometria) {
  const a = audiograma;
  if (!a) return "";
  const ptaOD = computePTA(a.od.va), ptaOE = computePTA(a.oe.va);
  if (ptaOD == null && ptaOE == null) return "Dados insuficientes para propor laudo — preencha ao menos 3 das 4 frequências (500, 1000, 2000 e 4000 Hz) por orelha na audiometria tonal.";
  const typeOD = classifyLossType(ptaOD, computePTA(a.od.vo));
  const typeOE = classifyLossType(ptaOE, computePTA(a.oe.vo));
  const gradeOD = classifyPTA(ptaOD), gradeOE = classifyPTA(ptaOE);
  const describeEar = (label, pta, grade, type) => {
    if (pta == null) return `${label}: dados insuficientes.`;
    if (grade === "Normal") return `${label}: limiares auditivos dentro dos padrões de normalidade (PTA ${pta} dB — OMS, 2021).`;
    return `${label}: limiares compatíveis com perda auditiva do tipo ${(type || "a definir").toLowerCase()}, grau ${grade.toLowerCase()} (PTA ${pta} dB — OMS, 2021).`;
  };
  const parts = [describeEar("Orelha direita", ptaOD, gradeOD, typeOD), describeEar("Orelha esquerda", ptaOE, gradeOE, typeOE)];
  if (gradeOD && gradeOE) parts.push(gradeOD === gradeOE && typeOD === typeOE ? "Quadro audiológico simétrico entre as orelhas." : "Quadro audiológico assimétrico entre as orelhas.");
  if (logoaudio && (logoaudio.iprf?.od?.mono || logoaudio.iprf?.oe?.mono)) parts.push(`Índice Percentual de Reconhecimento de Fala: OD ${logoaudio.iprf.od.mono || "—"}% e OE ${logoaudio.iprf.oe.mono || "—"}% (monossílabas).`);
  if (timpanometria && (timpanometria.classificacao?.od || timpanometria.classificacao?.oe)) parts.push(`Imitanciometria: curva timpanométrica ${timpanometria.classificacao.od || "—"} à direita e ${timpanometria.classificacao.oe || "—"} à esquerda.`);
  const allNormal = gradeOD === "Normal" && gradeOE === "Normal";
  parts.push(allNormal ? "Recomenda-se acompanhamento audiológico de rotina." : "Recomenda-se avaliação otorrinolaringológica e acompanhamento fonoaudiológico.");
  return parts.join(" ");
}

/* Masking advisor — estimates when interaural attenuation rules suggest
   masking may be needed, and a rough effective-masking range (plateau
   method). Teaching aid only: flags possibilities for the intern to
   confirm with the supervisor, not a clinical determination. */
const IA_AC = 40, ABG_MASK_THRESHOLD = 10, MIN_EFF_MASK_MARGIN = 10;
function computeMaskingAdvice(od, oe) {
  const advisories = [];
  const num = (v) => { const n = parseFloat(v); return isNaN(n) ? null : n; };
  FREQS_AUD.forEach((f) => {
    const acOD = num(od.va[f]), acOE = num(oe.va[f]), boOD = num(od.vo[f]), boOE = num(oe.vo[f]);
    if (acOD != null && boOE != null && acOD - boOE >= IA_AC) {
      advisories.push({ ear: "OD", cond: "va", freq: f, detail: `Via aérea OD ${acOD}dB − via óssea OE ${boOE}dB = ${acOD - boOE}dB (≥ ${IA_AC}dB de atenuação interaural)`,
        suggestion: `Mascarar OE — faixa efetiva estimada ${boOE + MIN_EFF_MASK_MARGIN}–${acOD + IA_AC - MIN_EFF_MASK_MARGIN} dB NA` });
    }
    if (acOE != null && boOD != null && acOE - boOD >= IA_AC) {
      advisories.push({ ear: "OE", cond: "va", freq: f, detail: `Via aérea OE ${acOE}dB − via óssea OD ${boOD}dB = ${acOE - boOD}dB (≥ ${IA_AC}dB de atenuação interaural)`,
        suggestion: `Mascarar OD — faixa efetiva estimada ${boOD + MIN_EFF_MASK_MARGIN}–${acOE + IA_AC - MIN_EFF_MASK_MARGIN} dB NA` });
    }
    if (acOD != null && boOD != null && acOD - boOD >= ABG_MASK_THRESHOLD) {
      advisories.push({ ear: "OD", cond: "vo", freq: f, detail: `Gap ar-ósseo OD de ${acOD - boOD}dB em ${f}Hz (≥ ${ABG_MASK_THRESHOLD}dB)`,
        suggestion: "Mascarar via óssea — atenuação interaural da via óssea é ~0dB" });
    }
    if (acOE != null && boOE != null && acOE - boOE >= ABG_MASK_THRESHOLD) {
      advisories.push({ ear: "OE", cond: "vo", freq: f, detail: `Gap ar-ósseo OE de ${acOE - boOE}dB em ${f}Hz (≥ ${ABG_MASK_THRESHOLD}dB)`,
        suggestion: "Mascarar via óssea — atenuação interaural da via óssea é ~0dB" });
    }
  });
  return advisories;
}
function MaskingAdvicePanel({ od, oe }) {
  const advisories = useMemo(() => computeMaskingAdvice(od, oe), [od, oe]);
  if (advisories.length === 0) return null;
  return (
    <div className="flex flex-col gap-1.5 p-3" style={{ background: T.warnSoft, borderRadius: 8, border: "1px solid #E3C6BC" }}>
      <div style={{ fontSize: 11.5, fontWeight: 600, color: T.warn }}>Possível necessidade de mascaramento (estimativa — confirme com a supervisão)</div>
      {advisories.map((a, i) => (
        <div key={i} style={{ fontSize: 11, color: T.ink }}>
          <span style={{ fontWeight: 600 }}>{a.ear} · {a.freq}Hz · {a.cond === "va" ? "via aérea" : "via óssea"}</span> — {a.detail}. <span style={{ color: T.inkSoft }}>{a.suggestion}.</span>
        </div>
      ))}
    </div>
  );
}

/* High-level combined clinical audiogram: both ears on one chart, ASHA-style
   symbols (OD red circle / OE blue X for air; OD red "<" / OE blue ">" for
   bone), severity bands, legend and auto-computed PTA per ear. */
function AudiogramChart({ od, oe }) {
  const W = 620, H = 400, padL = 54, padR = 24, padT = 18, padB = 46;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const xForFreq = (f) => padL + (Math.log2(f / 125) / Math.log2(8000 / 125)) * plotW;
  const yForDb = (db) => padT + ((db + 10) / 140) * plotH;
  const seriesFor = (ear) => ({
    va: FREQS_AUD.filter((f) => ear.va[f] !== "" && ear.va[f] != null && !isNaN(ear.va[f])).map((f) => [xForFreq(f), yForDb(Number(ear.va[f]))]),
    vo: FREQS_AUD.filter((f) => ear.vo[f] !== "" && ear.vo[f] != null && !isNaN(ear.vo[f])).map((f) => [xForFreq(f), yForDb(Number(ear.vo[f]))]),
  });
  const odS = seriesFor(od), oeS = seriesFor(oe);
  const dbLines = [-10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120];
  const ptaOD = computePTA(od.va), ptaOE = computePTA(oe.va);

  return (
    <div className="flex flex-col gap-2">
      <div className="flex flex-wrap gap-4 items-center">
        <div style={{ fontSize: 12.5, fontFamily: "'IBM Plex Sans', sans-serif" }}>
          <span style={{ color: OD_COLOR, fontWeight: 600 }}>OD</span>{" "}
          <span style={{ color: T.inkSoft }}>{ptaOD != null ? `PTA (500–4k) ${ptaOD} dB — ${classifyPTA(ptaOD)}` : "PTA — dados insuficientes"}</span>
        </div>
        <div style={{ fontSize: 12.5, fontFamily: "'IBM Plex Sans', sans-serif" }}>
          <span style={{ color: OE_COLOR, fontWeight: 600 }}>OE</span>{" "}
          <span style={{ color: T.inkSoft }}>{ptaOE != null ? `PTA (500–4k) ${ptaOE} dB — ${classifyPTA(ptaOE)}` : "PTA — dados insuficientes"}</span>
        </div>
        <span style={{ fontSize: 10, color: T.inkFaint, fontStyle: "italic" }}>Classificação: OMS, World Report on Hearing (2021)</span>
      </div>
      <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: 620, background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 8 }}>
        {SEVERITY_BANDS.map((b) => (
          <rect key={b.label} x={padL} width={plotW} y={yForDb(b.to)} height={yForDb(b.from) - yForDb(b.to)} fill={b.color} opacity="0.07" />
        ))}
        {dbLines.map((db) => (
          <g key={db}>
            <line x1={padL} x2={W - padR} y1={yForDb(db)} y2={yForDb(db)} stroke={db % 20 === 0 ? T.line : T.lineSoft} strokeWidth={db % 20 === 0 ? 1 : 0.6} />
            <text x={padL - 6} y={yForDb(db) + 3} fontSize="9" textAnchor="end" fill={T.inkFaint} fontFamily="'IBM Plex Sans', sans-serif">{db}</text>
          </g>
        ))}
        {FREQS_AUD.map((f) => (
          <line key={f} x1={xForFreq(f)} x2={xForFreq(f)} y1={padT} y2={H - padB} stroke={T.lineSoft} strokeWidth="0.6" />
        ))}
        {[125, 250, 500, 1000, 2000, 4000, 8000].map((f) => (
          <text key={f} x={xForFreq(f)} y={H - padB + 16} fontSize="9.5" textAnchor="middle" fill={T.inkSoft} fontFamily="'IBM Plex Sans', sans-serif">{f >= 1000 ? f / 1000 + "k" : f}</text>
        ))}
        <text x={2} y={padT + plotH / 2} fontSize="9.5" fill={T.inkFaint} fontFamily="'IBM Plex Sans', sans-serif" transform={`rotate(-90, 12, ${padT + plotH / 2})`}>Nível de audição (dB NA)</text>
        <text x={padL + plotW / 2} y={H - 4} fontSize="9.5" textAnchor="middle" fill={T.inkFaint} fontFamily="'IBM Plex Sans', sans-serif">Frequência (Hz)</text>

        {odS.va.length > 1 && <polyline points={odS.va.map((p) => p.join(",")).join(" ")} fill="none" stroke={OD_COLOR} strokeWidth="1.5" />}
        {oeS.va.length > 1 && <polyline points={oeS.va.map((p) => p.join(",")).join(" ")} fill="none" stroke={OE_COLOR} strokeWidth="1.5" strokeDasharray="5,3" />}
        {odS.va.map(([x, y], i) => <circle key={"odva" + i} cx={x} cy={y} r="5" fill="white" stroke={OD_COLOR} strokeWidth="1.6" />)}
        {oeS.va.map(([x, y], i) => (
          <g key={"oeva" + i}>
            <line x1={x - 4.5} y1={y - 4.5} x2={x + 4.5} y2={y + 4.5} stroke={OE_COLOR} strokeWidth="1.6" />
            <line x1={x - 4.5} y1={y + 4.5} x2={x + 4.5} y2={y - 4.5} stroke={OE_COLOR} strokeWidth="1.6" />
          </g>
        ))}
        {odS.vo.map(([x, y], i) => <path key={"odvo" + i} d={`M${x + 4},${y - 4} L${x - 4},${y} L${x + 4},${y + 4}`} fill="none" stroke={OD_COLOR} strokeWidth="1.6" />)}
        {oeS.vo.map(([x, y], i) => <path key={"oevo" + i} d={`M${x - 4},${y - 4} L${x + 4},${y} L${x - 4},${y + 4}`} fill="none" stroke={OE_COLOR} strokeWidth="1.6" />)}
      </svg>
      <div className="flex flex-wrap gap-x-5 gap-y-1.5" style={{ fontSize: 11, color: T.inkSoft, fontFamily: "'IBM Plex Sans', sans-serif" }}>
        <span className="flex items-center gap-1"><svg width="14" height="14"><circle cx="7" cy="7" r="5" fill="white" stroke={OD_COLOR} strokeWidth="1.6" /></svg> OD — via aérea</span>
        <span className="flex items-center gap-1"><svg width="14" height="14"><path d="M11,3 L3,7 L11,11" fill="none" stroke={OD_COLOR} strokeWidth="1.6" /></svg> OD — via óssea</span>
        <span className="flex items-center gap-1"><svg width="14" height="14"><line x1="3" y1="3" x2="11" y2="11" stroke={OE_COLOR} strokeWidth="1.6" /><line x1="3" y1="11" x2="11" y2="3" stroke={OE_COLOR} strokeWidth="1.6" /></svg> OE — via aérea</span>
        <span className="flex items-center gap-1"><svg width="14" height="14"><path d="M3,3 L11,7 L3,11" fill="none" stroke={OE_COLOR} strokeWidth="1.6" /></svg> OE — via óssea</span>
      </div>
    </div>
  );
}
/* Actual masking log — records, per threshold obtained with masking, the
   initial masking level applied and which ear received the masking noise
   (student-entered, distinct from the automatic estimate above). */
const EAR_LABEL = { od: "OD", oe: "OE" };
function MaskingLogEditor({ rows, onChange, disabled }) {
  const list = rows || [];
  const add = () => onChange([...list, { id: genId(), orelhaTestada: "od", via: "va", freq: "1000", orelhaMascarada: "oe", nivelInicial: "", nivelEfetivo: "" }]);
  const update = (id, field, v) => onChange(list.map((m) => m.id === id ? { ...m, [field]: v, ...(field === "orelhaTestada" ? { orelhaMascarada: v === "od" ? "oe" : "od" } : {}) } : m));
  const remove = (id) => onChange(list.filter((m) => m.id !== id));
  const selectStyle = { ...cellInput, width: "auto", padding: "3px 4px" };
  return (
    <SubBlock title="Mascaramento aplicado — orelha mascarada e nível inicial">
      <div className="flex flex-col gap-2">
        {list.length === 0 && <div style={{ fontSize: 11.5, color: T.inkFaint }}>Nenhum registro ainda. Adicione um para cada limiar obtido com mascaramento.</div>}
        {list.map((m) => (
          <div key={m.id} className="flex flex-wrap items-center gap-2 p-2" style={{ background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 6 }}>
            <label className="flex items-center gap-1" style={{ fontSize: 11, color: T.inkSoft }}>Orelha testada
              <select value={m.orelhaTestada} disabled={disabled} onChange={(e) => update(m.id, "orelhaTestada", e.target.value)} style={selectStyle}>
                <option value="od">OD</option><option value="oe">OE</option>
              </select>
            </label>
            <label className="flex items-center gap-1" style={{ fontSize: 11, color: T.inkSoft }}>Via
              <select value={m.via} disabled={disabled} onChange={(e) => update(m.id, "via", e.target.value)} style={selectStyle}>
                <option value="va">Aérea</option><option value="vo">Óssea</option>
              </select>
            </label>
            <label className="flex items-center gap-1" style={{ fontSize: 11, color: T.inkSoft }}>Freq.
              <select value={m.freq} disabled={disabled} onChange={(e) => update(m.id, "freq", e.target.value)} style={selectStyle}>
                {FREQS_AUD.map((f) => <option key={f} value={f}>{f} Hz</option>)}
              </select>
            </label>
            <label className="flex items-center gap-1" style={{ fontSize: 11, color: T.warn, fontWeight: 600 }}>Orelha mascarada
              <select value={m.orelhaMascarada} disabled={disabled} onChange={(e) => update(m.id, "orelhaMascarada", e.target.value)} style={selectStyle}>
                <option value="od">OD</option><option value="oe">OE</option>
              </select>
            </label>
            <label className="flex items-center gap-1" style={{ fontSize: 11, color: T.inkSoft }}>Mascaramento inicial (dB)
              <input value={m.nivelInicial} disabled={disabled} onChange={(e) => update(m.id, "nivelInicial", e.target.value)} style={{ ...cellInput, width: 56 }} />
            </label>
            <label className="flex items-center gap-1" style={{ fontSize: 11, color: T.inkSoft }}>Efetivo (dB, opcional)
              <input value={m.nivelEfetivo} disabled={disabled} onChange={(e) => update(m.id, "nivelEfetivo", e.target.value)} style={{ ...cellInput, width: 56 }} />
            </label>
            {!disabled && <button onClick={() => remove(m.id)} style={{ color: T.warn, marginLeft: "auto" }}><Trash2 size={14} /></button>}
          </div>
        ))}
        {!disabled && (
          <button onClick={add} className="flex items-center gap-1.5 self-start" style={{ fontSize: 12, color: T.accent, fontWeight: 500 }}>
            <Plus size={14} /> Adicionar registro de mascaramento
          </button>
        )}
      </div>
    </SubBlock>
  );
}
function maskingLogText(rows) {
  if (!rows || rows.length === 0) return "Nenhum registro";
  return rows.map((m) => `${EAR_LABEL[m.orelhaTestada]} ${m.freq}Hz ${m.via === "va" ? "VA" : "VO"} — mascarado em ${EAR_LABEL[m.orelhaMascarada]}, inicial ${m.nivelInicial || "—"}dB${m.nivelEfetivo ? `, efetivo ${m.nivelEfetivo}dB` : ""}`).join(" · ");
}

/* ---------------------------------------------------------------------
   Protocolo Aplicado — módulo genérico, presente em todas as áreas.
   Registra qual protocolo formal foi aplicado na sessão, com uma lista
   de itens/escores flexível. Os modelos abaixo pré-carregam os itens
   com a estrutura de pontuação/resumo dos instrumentos padronizados
   (não reproduzem os estímulos/itens autorais do teste em si).
------------------------------------------------------------------------ */
const blankProtocoloAplicado = () => ({ nome: "", dataAplicacao: "", referencia: "", escoreGeral: "", itens: [], obs: "" });

const mkItens = (names) => names.map((n) => ({ id: genId(), item: n, resultado: "" }));

const PROTOCOL_LIBRARY = [
  {
    group: "Motricidade Orofacial",
    key: "mbgr",
    nome: "MBGR — Exame Miofuncional Orofacial",
    referencia: "Marchesan IQ, Berretin-Felix G, Genaro KF, Rehder MI",
    itens: mkItens([
      "1. Postura Corporal (0–7)",
      "2. Exame Extraoral (0–17)",
      "  2a. Face (0–5)", "  2b. Lábios (0–10)", "  2c. Masseter (0–2)",
      "3. Exame Intraoral (0–57)",
      "  3a. Lábios (0–5)", "  3b. Língua (0–17)", "  3c. Bochechas (0–8)",
      "  3d. Palato (0–8)", "  3e. Tonsilas Palatinas (0–4)", "  3f. Dentes (0–5)", "  3g. Oclusão (0–10)",
      "4. Mobilidade (0–54)",
      "  4a. Lábios (0–16)", "  4b. Língua (0–16)", "  4c. Véu Palatino (0–4)", "  4d. Mandíbula (0–18)",
      "5. Dor à Palpação (0–10)",
      "6. Tônus (0–6)",
      "  6a. Lábios (0–2)", "  6b. Mento (0–1)", "  6c. Língua (0–1)", "  6d. Bochechas (0–2)",
      "7. Funções Orofaciais (0–103)",
      "  7a. Respiração (0–7)", "  7b. Mastigação (0–10)", "  7c. Deglutição (0–39)", "  7d. Fala (0–44)", "  7e. Voz (0–3)",
    ]),
  },
  {
    group: "Linguagem Infantil / Fala",
    key: "abfw_fonologia",
    nome: "ABFW — Fonologia (Quadro Resumo)",
    referencia: "Wertzner HF, in ABFW (Andrade, Befi-Lopes, Fernandes, Wertzner), Pró-Fono, 2004",
    itens: mkItens([
      "Redução de sílaba", "Harmonia consonantal", "Plosivação de fricativas",
      "Posteriorização para velar", "Posteriorização para palatal", "Frontalização de velares",
      "Frontalização de palatal", "Simplificação de líquida", "Simplificação do encontro consonantal",
      "Simplificação da consoante final", "Sonorização de plosivas", "Sonorização de fricativas",
      "Ensurdecimento de plosivas", "Ensurdecimento de fricativas", "Outros",
    ]),
  },
  {
    group: "Linguagem Infantil / Fala",
    key: "abfw_vocabulario",
    nome: "ABFW — Vocabulário (Tabela Síntese)",
    referencia: "Befi-Lopes DM, in ABFW (Andrade, Befi-Lopes, Fernandes, Wertzner), Pró-Fono, 2004",
    itens: mkItens([
      "Vestuário", "Animais", "Alimentos", "Meios de transporte", "Móveis e utensílios",
      "Profissões", "Locais", "Formas e cores", "Brinquedos e instrumentos musicais",
    ]),
  },
  {
    group: "Linguagem Infantil / Fala",
    key: "abfw_fluencia",
    nome: "ABFW — Fluência (Protocolo de Avaliação)",
    referencia: "Andrade CRF, in ABFW (Andrade, Befi-Lopes, Fernandes, Wertzner), Pró-Fono, 2004",
    itens: mkItens([
      "Hesitação", "Interjeição", "Revisão", "Palavra não terminada", "Repetição de palavras",
      "Repetição de segmentos", "Repetição de frases", "Total disfluências comuns",
      "Repetição de sílabas", "Repetição de sons", "Prolongamento", "Bloqueio", "Pausa",
      "Intrusão de sons ou segmentos", "Total disfluências gagas",
      "Fluxo de palavras por minuto", "Fluxo de sílabas por minuto",
      "% Descontinuidade de fala", "% Disfluências gagas",
    ]),
  },
  {
    group: "Linguagem Infantil / Fala",
    key: "abfw_pragmatica",
    nome: "ABFW — Pragmática (Ficha-Síntese)",
    referencia: "Fernandes FDM, in ABFW (Andrade, Befi-Lopes, Fernandes, Wertzner), Pró-Fono, 2004",
    itens: mkItens([
      "Total de atos comunicativos", "Atos por minuto", "% do total",
      "PO", "RO", "EX", "EP", "PR", "PE", "NA",
      "PS", "C", "NF", "PA", "E", "JC", "RE",
      "PI", "N", "XP", "PC", "AR", "J",
    ]),
  },
  {
    group: "Linguagem Infantil / Fala",
    key: "adl2",
    nome: "ADL2 — Avaliação do Desenvolvimento da Linguagem",
    referencia: "Menezes ML, PhD — ADL2, Protocolo de Aplicação e Pontuação",
    itens: mkItens([
      "1a a 1a5m — LC/LE", "1a6m a 1a11m — LC/LE", "2a a 2a5m — LC/LE", "2a6m a 2a11m — LC/LE",
      "3a a 3a5m — LC/LE", "3a6m a 3a11m — LC/LE", "4a a 4a5m — LC/LE", "4a6m a 4a11m — LC/LE",
      "5a a 5a5m — LC/LE", "5a6m a 5a11m — LC/LE", "6a a 6a5m — LC/LE", "6a6m a 6a11m — LC/LE",
      "Linguagem Receptiva — Última tarefa correta", "Linguagem Receptiva — Menos respostas incorretas",
      "Linguagem Receptiva — Escore Bruto", "Linguagem Receptiva — Escore Padrão",
      "Linguagem Expressiva — Última tarefa correta", "Linguagem Expressiva — Menos respostas incorretas",
      "Linguagem Expressiva — Escore Bruto", "Linguagem Expressiva — Escore Padrão",
      "Linguagem Global — Escore Bruto", "Linguagem Global — Escore Padrão",
    ]),
  },
  {
    group: "Linguagem Infantil / Fala",
    key: "adl2_fonologia",
    nome: "ADL2 — Observação da Aquisição Fonológica e do Vocabulário",
    referencia: "Menezes ML, PhD — complementar à ADL2 (não é um teste fonoaudiológico padronizado)",
    itens: mkItens([
      "3a a 3a5m (itens 1–19)", "3a6m a 3a11m (itens 20–24)", "4a a 4a5m (itens 25–28)",
      "4a6m a 4a11m (itens 29–32)", "5a a 5a11m (itens 33–40)",
    ]),
  },
  {
    group: "Linguagem Infantil / Fala",
    key: "proc",
    nome: "PROC — Protocolo de Observação Comportamental",
    referencia: "PROC — Protocolo de Observação Comportamental",
    itens: mkItens([
      "1. Habilidades comunicativas expressivas (máx. 70)",
      "  1a. Habilidades dialógicas/conversacionais", "  1b. Funções comunicativas",
      "2. Compreensão da linguagem verbal (máx. 60)",
      "3. Aspectos do desenvolvimento cognitivo (máx. 70)",
      "  3a. Formas de manipulação dos objetos",
      "TOTAL (máx. 200)",
    ]),
  },
];

/* ---------------------------------------------------------------------
   Cálculo automático do "Protocolo Aplicado": soma apenas os itens de
   nível principal (não indentados, e que não sejam a própria linha de
   "TOTAL") para não contar domínios e subdomínios em duplicidade. O
   valor máximo é extraído do próprio rótulo do item quando presente,
   no formato "(0–N)" ou "(máx. N)".
------------------------------------------------------------------------ */
function isSubItemLabel(label) { return /^\s{2}/.test(label || ""); }
function isTotalItemLabel(label) { return /total/i.test(label || ""); }
function parseMaxFromLabel(label) {
  const m = (label || "").match(/\(0\s*[–-]\s*(\d+(?:[.,]\d+)?)\)/) || (label || "").match(/máx\.?\s*(\d+(?:[.,]\d+)?)/i);
  return m ? parseFloat(m[1].replace(",", ".")) : null;
}
function computeProtocoloTotais(itens) {
  const summable = (itens || []).filter((it) => !isSubItemLabel(it.item) && !isTotalItemLabel(it.item) && it.resultado !== "" && !isNaN(parseFloat(it.resultado)));
  if (summable.length === 0) return null;
  const total = summable.reduce((a, it) => a + parseFloat(it.resultado), 0);
  const maxes = summable.map((it) => parseMaxFromLabel(it.item));
  const max = maxes.every((m) => m != null) ? maxes.reduce((a, m) => a + m, 0) : null;
  return { total, max, pct: max ? Math.round((total / max) * 100) : null, count: summable.length };
}
function ProtocoloChart({ itens }) {
  const rows = (itens || [])
    .filter((it) => !isSubItemLabel(it.item) && it.item && it.resultado !== "" && !isNaN(parseFloat(it.resultado)))
    .map((it) => ({ label: it.item.length > 34 ? it.item.slice(0, 33) + "…" : it.item, value: parseFloat(it.resultado), max: parseMaxFromLabel(it.item) }));
  if (rows.length === 0) return null;
  const W = 480, rowH = 22, padL = 190, padR = 46, padT = 6;
  const H = padT + rows.length * rowH + 4;
  const globalMax = Math.max(...rows.map((r) => r.max || r.value || 1), 1);
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: 480, background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 6 }}>
      {rows.map((r, i) => {
        const y = padT + i * rowH;
        const scaleMax = r.max || globalMax;
        const barW = Math.max(2, ((W - padL - padR) * Math.min(r.value, scaleMax)) / scaleMax);
        const over = r.value > scaleMax;
        return (
          <g key={i}>
            <text x={padL - 6} y={y + rowH / 2 + 3} fontSize="9" textAnchor="end" fill={T.inkSoft}>{r.label}</text>
            <rect x={padL} y={y + 4} width={W - padL - padR} height={rowH - 9} fill={T.lineSoft} rx="2" />
            <rect x={padL} y={y + 4} width={barW} height={rowH - 9} fill={over ? T.warn : T.accent} rx="2" />
            <text x={padL + Math.max(barW, 14) + 4} y={y + rowH / 2 + 3} fontSize="9" fill={T.inkFaint}>{r.value}{r.max != null ? ` / ${r.max}` : ""}</text>
          </g>
        );
      })}
    </svg>
  );
}
function gerarSinteseProtocolo(value) {
  const t = computeProtocoloTotais(value.itens);
  if (!t) return "Preencha os resultados numéricos dos itens para gerar uma síntese automática.";
  const topItens = (value.itens || []).filter((it) => !isSubItemLabel(it.item) && !isTotalItemLabel(it.item) && it.resultado !== "" && !isNaN(parseFloat(it.resultado)));
  const pior = topItens.length ? topItens.reduce((a, b) => {
    const pa = parseMaxFromLabel(a.item), pb = parseMaxFromLabel(b.item);
    const ra = pa ? parseFloat(a.resultado) / pa : parseFloat(a.resultado);
    const rb = pb ? parseFloat(b.resultado) / pb : parseFloat(b.resultado);
    return rb > ra ? b : a;
  }) : null;
  const parts = [];
  parts.push(`Na aplicação do protocolo ${value.nome || "informado"}${value.dataAplicacao ? " em " + value.dataAplicacao : ""}, obteve-se escore total de ${t.total}${t.max ? ` em ${t.max} pontos possíveis (${t.pct}%)` : ""}.`);
  if (pior) parts.push(`O domínio com maior pontuação/alteração foi "${pior.item.trim()}" (${pior.resultado}${parseMaxFromLabel(pior.item) ? "/" + parseMaxFromLabel(pior.item) : ""}).`);
  parts.push("Recomenda-se correlacionar este resultado com a avaliação clínica e, quando aplicável, comparar com reavaliações futuras para acompanhar a evolução.");
  return parts.join(" ");
}

function ProtocoloAplicadoEditor({ value, onChange, disabled }) {
  const applyTemplate = (key) => {
    const tpl = PROTOCOL_LIBRARY.find((t) => t.key === key);
    if (!tpl) return;
    if ((value.itens || []).length > 0 && !window.confirm("Isso substitui os itens já preenchidos por este modelo. Continuar?")) return;
    onChange("nome", tpl.nome);
    onChange("referencia", tpl.referencia);
    onChange("itens", mkItens(tpl.itens.map((it) => it.item)));
  };
  const itens = value.itens || [];
  const addItem = () => onChange("itens", [...itens, { id: genId(), item: "", resultado: "" }]);
  const updateItem = (id, field, v) => onChange("itens", itens.map((it) => (it.id === id ? { ...it, [field]: v } : it)));
  const removeItem = (id) => onChange("itens", itens.filter((it) => it.id !== id));
  const totais = computeProtocoloTotais(itens);
  const calcularEscore = () => { if (totais) onChange("escoreGeral", `${totais.total}${totais.max ? ` / ${totais.max} (${totais.pct}%)` : ""}`); };
  return (
    <div className="flex flex-col gap-3">
      {!disabled && (
        <Field label="Usar um modelo de protocolo pronto (opcional)">
          <select defaultValue="" onChange={(e) => { if (e.target.value) applyTemplate(e.target.value); e.target.value = ""; }} style={inputBase}>
            <option value="">Selecionar modelo…</option>
            {Object.entries(_.groupBy(PROTOCOL_LIBRARY, "group")).map(([group, tpls]) => (
              <optgroup key={group} label={group}>
                {tpls.map((t) => <option key={t.key} value={t.key}>{t.nome}</option>)}
              </optgroup>
            ))}
          </select>
        </Field>
      )}
      <div className="grid sm:grid-cols-2 gap-3">
        <Field label="Nome do protocolo"><TextInput value={value.nome} disabled={disabled} onChange={(v) => onChange("nome", v)} placeholder="Ex.: (a definir quando o protocolo for informado)" /></Field>
        <Field label="Data de aplicação"><TextInput type="date" value={value.dataAplicacao} disabled={disabled} onChange={(v) => onChange("dataAplicacao", v)} /></Field>
        <Field label="Referência / versão"><TextInput value={value.referencia} disabled={disabled} onChange={(v) => onChange("referencia", v)} /></Field>
        <Field label="Escore geral / resultado">
          <div className="flex gap-2">
            <TextInput value={value.escoreGeral} disabled={disabled} onChange={(v) => onChange("escoreGeral", v)} />
            {!disabled && totais && <button onClick={calcularEscore} title="Somar automaticamente os itens de nível principal" style={{ fontSize: 11, color: T.accent, whiteSpace: "nowrap" }}>Calcular</button>}
          </div>
        </Field>
      </div>
      <SubBlock title="Itens avaliados">
        <div className="flex flex-col gap-2">
          {itens.length === 0 && <div style={{ fontSize: 11.5, color: T.inkFaint }}>Nenhum item adicionado.</div>}
          {itens.map((it) => (
            <div key={it.id} className="flex flex-wrap items-center gap-2 p-2" style={{ background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 6 }}>
              <input value={it.item} disabled={disabled} onChange={(e) => updateItem(it.id, "item", e.target.value)} placeholder="Item / domínio avaliado" style={{ ...cellInput, width: 220, textAlign: "left" }} />
              <input value={it.resultado} disabled={disabled} onChange={(e) => updateItem(it.id, "resultado", e.target.value)} placeholder="Resultado / escore" style={{ ...cellInput, width: 160, textAlign: "left" }} />
              {!disabled && <button onClick={() => removeItem(it.id)} style={{ color: T.warn, marginLeft: "auto" }}><Trash2 size={14} /></button>}
            </div>
          ))}
          {!disabled && (
            <button onClick={addItem} className="flex items-center gap-1.5 self-start" style={{ fontSize: 12, color: T.accent, fontWeight: 500 }}>
              <Plus size={14} /> Adicionar item
            </button>
          )}
        </div>
      </SubBlock>
      {totais && (
        <div className="flex flex-col gap-1.5 p-3" style={{ background: T.accentSoft, borderRadius: 8 }}>
          <div style={{ fontSize: 12, fontWeight: 600, color: T.accentDeep }}>
            Escore calculado (itens de nível principal): {totais.total}{totais.max ? ` / ${totais.max} (${totais.pct}%)` : ""}
          </div>
          <ProtocoloChart itens={itens} />
        </div>
      )}
      <Field label="Observações / interpretação">
        <TextArea rows={3} value={value.obs} disabled={disabled} onChange={(v) => onChange("obs", v)} />
      </Field>
      {!disabled && (
        <button onClick={() => onChange("obs", gerarSinteseProtocolo(value))} className="flex items-center gap-1.5 self-start" style={{ fontSize: 12, color: T.accent, fontWeight: 500 }}>
          <FileBarChart size={14} /> Gerar síntese automática
        </button>
      )}
    </div>
  );
}
function ProtocoloAplicadoReadView({ value }) {
  if (!value?.nome && (!value?.itens || value.itens.length === 0)) return <ReadRow value="Nenhum protocolo registrado nesta sessão." />;
  const totais = computeProtocoloTotais(value.itens);
  return (
    <div className="flex flex-col gap-2">
      <ReadRow label="Protocolo" value={value.nome} />
      <ReadRow label="Data de aplicação" value={value.dataAplicacao} />
      <ReadRow label="Referência / versão" value={value.referencia} />
      <ReadRow label="Escore geral" value={value.escoreGeral} />
      {totais && <ReadRow label="Escore calculado" value={`${totais.total}${totais.max ? ` / ${totais.max} (${totais.pct}%)` : ""}`} />}
      {(value.itens || []).length > 0 && (
        <ReadRow label="Itens avaliados" value={value.itens.map((it) => `${it.item || "—"}: ${it.resultado || "—"}`).join(" · ")} />
      )}
      {totais && <div className="my-1"><ProtocoloChart itens={value.itens} /></div>}
      <ReadRow label="Observações" value={value.obs} />
    </div>
  );
}
function protocoloAplicadoTextLines(value) {
  if (!value?.nome && (!value?.itens || value.itens.length === 0)) return [["Protocolo aplicado", "Nenhum registrado"]];
  const totais = computeProtocoloTotais(value.itens);
  return [
    ["Protocolo", value.nome || "—"],
    ["Data de aplicação", value.dataAplicacao || "—"],
    ["Referência / versão", value.referencia || "—"],
    ["Escore geral", value.escoreGeral || "—"],
    ...(totais ? [["Escore calculado", `${totais.total}${totais.max ? ` / ${totais.max} (${totais.pct}%)` : ""}`]] : []),
    ["Itens avaliados", (value.itens || []).map((it) => `${it.item || "—"}: ${it.resultado || "—"}`).join(" · ") || "—"],
    ["Observações", value.obs || "—"],
  ];
}

// Anexa a seção "Protocolo Aplicado" ao final de cada área — genérica por
// enquanto; será detalhada por área assim que os protocolos forem enviados.
Object.values(AREAS).forEach((area) => {
  area.sections.push({
    title: "Protocolo Aplicado",
    children: [ cst("protocoloAplicado", blankProtocoloAplicado, ProtocoloAplicadoEditor, ProtocoloAplicadoReadView, protocoloAplicadoTextLines) ],
  });
});

function AudiogramEditor({ value, onChange, disabled }) {
  const setEar = (ear, cond, freq, v) => onChange(`${ear}.${cond}.${freq}`, v);
  const Table = ({ ear, earLabel }) => (
    <div className="overflow-x-auto">
      <table style={{ borderCollapse: "collapse", width: "100%" }}>
        <thead>
          <tr>
            <th style={{ fontSize: 11, color: T.inkSoft, textAlign: "left", padding: "2px 4px" }}>{earLabel}</th>
            {FREQS_AUD.map((f) => <th key={f} style={{ fontSize: 10, color: T.inkFaint, padding: "2px 3px" }}>{f}</th>)}
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style={{ fontSize: 11, color: T.inkSoft, padding: "2px 4px" }}>Via aérea</td>
            {FREQS_AUD.map((f) => <td key={f} style={{ padding: 1 }}><Cell value={value[ear].va[f]} disabled={disabled} onChange={(v) => setEar(ear, "va", f, v)} /></td>)}
          </tr>
          <tr>
            <td style={{ fontSize: 11, color: T.inkSoft, padding: "2px 4px" }}>Via óssea</td>
            {FREQS_AUD.map((f) => <td key={f} style={{ padding: 1 }}><Cell value={value[ear].vo[f]} disabled={disabled} onChange={(v) => setEar(ear, "vo", f, v)} /></td>)}
          </tr>
        </tbody>
      </table>
    </div>
  );
  return (
    <div className="flex flex-col gap-4">
      <Table ear="od" earLabel="Orelha direita (dB NA)" />
      <Table ear="oe" earLabel="Orelha esquerda (dB NA)" />
      <AudiogramChart od={value.od} oe={value.oe} />
      <LossTypeSummary od={value.od} oe={value.oe} />
      <MaskingAdvicePanel od={value.od} oe={value.oe} />
      <MaskingLogEditor rows={value.mascaramentos} disabled={disabled} onChange={(rows) => onChange("mascaramentos", rows)} />
      <SubBlock title="Laudo Audiológico">
        {!disabled && (
          <button onClick={() => onChange("laudo.obs", generateLaudoText(value))} className="flex items-center gap-1.5 self-start" style={{ fontSize: 12, color: T.accent, fontWeight: 500 }}>
            <FileBarChart size={14} /> Propor laudo (rascunho a partir dos dados)
          </button>
        )}
        <div style={{ fontSize: 10.5, color: T.inkFaint, marginTop: -4 }}>Gera um rascunho a partir do PTA, tipo de perda e demais exames preenchidos — revise e ajuste antes de assinar.</div>
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="OD"><TextInput value={value.laudo.od} disabled={disabled} onChange={(v) => onChange("laudo.od", v)} /></Field>
          <Field label="OE"><TextInput value={value.laudo.oe} disabled={disabled} onChange={(v) => onChange("laudo.oe", v)} /></Field>
        </div>
        <Field label="OBS"><TextArea rows={4} value={value.laudo.obs} disabled={disabled} onChange={(v) => onChange("laudo.obs", v)} /></Field>
        <Field label="Referência"><TextInput value={value.laudo.referencia} disabled={disabled} onChange={(v) => onChange("laudo.referencia", v)} /></Field>
      </SubBlock>
    </div>
  );
}
function AudiogramReadView({ value }) {
  const fmt = (ear, cond) => FREQS_AUD.map((f) => value[ear][cond][f]).some(Boolean)
    ? FREQS_AUD.filter((f) => value[ear][cond][f]).map((f) => `${f}Hz: ${value[ear][cond][f]}dB`).join(" · ") : "";
  return (
    <div className="flex flex-col gap-2">
      <ReadRow label="OD — via aérea" value={fmt("od", "va")} />
      <ReadRow label="OD — via óssea" value={fmt("od", "vo")} />
      <ReadRow label="OE — via aérea" value={fmt("oe", "va")} />
      <ReadRow label="OE — via óssea" value={fmt("oe", "vo")} />
      <div className="my-1"><AudiogramChart od={value.od} oe={value.oe} /></div>
      <LossTypeSummary od={value.od} oe={value.oe} />
      <MaskingAdvicePanel od={value.od} oe={value.oe} />
      <ReadRow label="Mascaramento aplicado" value={maskingLogText(value.mascaramentos)} />
      <ReadRow label="Laudo — OD" value={value.laudo.od} />
      <ReadRow label="Laudo — OE" value={value.laudo.oe} />
      <ReadRow label="Laudo — OBS" value={value.laudo.obs} />
      <ReadRow label="Referência" value={value.laudo.referencia} />
    </div>
  );
}

/* IPRF grouped bar chart — % de acertos monossílabas/dissílabas por orelha */
function IPRFBarChart({ iprf }) {
  const W = 280, H = 150, padL = 28, padR = 8, padT = 10, padB = 22;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const groups = [
    { label: "OD", color: OD_COLOR, mono: Number(iprf.od.mono) || 0, dissi: Number(iprf.od.dissi) || 0 },
    { label: "OE", color: OE_COLOR, mono: Number(iprf.oe.mono) || 0, dissi: Number(iprf.oe.dissi) || 0 },
  ];
  const barW = 22, gap = 10, groupW = barW * 2 + gap;
  const yForPct = (p) => padT + plotH - (Math.min(100, Math.max(0, p)) / 100) * plotH;
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: 280, background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 6 }}>
      {[0, 25, 50, 75, 100].map((p) => (
        <g key={p}>
          <line x1={padL} x2={W - padR} y1={yForPct(p)} y2={yForPct(p)} stroke={T.lineSoft} strokeWidth="0.6" />
          <text x={padL - 4} y={yForPct(p) + 3} fontSize="7.5" textAnchor="end" fill={T.inkFaint}>{p}</text>
        </g>
      ))}
      {groups.map((g, gi) => {
        const gx = padL + 16 + gi * (groupW + 24);
        return (
          <g key={g.label}>
            <rect x={gx} y={yForPct(g.mono)} width={barW} height={plotH + padT - yForPct(g.mono)} fill={g.color} opacity="0.85" />
            <rect x={gx + barW + gap - barW} y={yForPct(g.dissi)} width={barW} height={plotH + padT - yForPct(g.dissi)} fill={g.color} opacity="0.35" />
            <text x={gx + barW} y={H - 8} fontSize="8.5" textAnchor="middle" fill={T.inkSoft}>{g.label}</text>
          </g>
        );
      })}
      <text x={W - 4} y={12} fontSize="7" textAnchor="end" fill={T.inkFaint}>% acertos</text>
    </svg>
  );
}

function LogoaudioEditor({ value, onChange, disabled }) {
  const F = ({ v, path, w }) => <input value={v || ""} disabled={disabled} onChange={(e) => onChange(path, e.target.value)} style={{ ...cellInput, width: w || 56 }} />;
  return (
    <div className="grid sm:grid-cols-2 gap-4">
      <div>
        <div style={{ fontSize: 12, fontWeight: 600, color: T.accentDeep, marginBottom: 4 }}>SRT/LRF · LDF</div>
        <table style={{ borderCollapse: "collapse", width: "100%" }}><thead><tr>
          <th></th><th style={{ fontSize: 10, color: T.inkFaint }}>SRT/LRF (dB)</th><th style={{ fontSize: 10, color: T.inkFaint }}>LDF (dB)</th><th style={{ fontSize: 10, color: T.inkFaint }}>Masc. (dB SN)</th>
        </tr></thead><tbody>
          <tr><td style={{ fontSize: 11, color: "#C0392B" }}>O.D.</td><td><F v={value.od.srt} path="od.srt" /></td><td><F v={value.od.ldf} path="od.ldf" /></td><td><F v={value.od.masc} path="od.masc" /></td></tr>
          <tr><td style={{ fontSize: 11, color: "#2C6E62" }}>O.E.</td><td><F v={value.oe.srt} path="oe.srt" /></td><td><F v={value.oe.ldf} path="oe.ldf" /></td><td><F v={value.oe.masc} path="oe.masc" /></td></tr>
        </tbody></table>
      </div>
      <div>
        <div style={{ fontSize: 12, fontWeight: 600, color: T.accentDeep, marginBottom: 4 }}>IPRF</div>
        <table style={{ borderCollapse: "collapse", width: "100%" }}><thead><tr>
          <th></th><th style={{ fontSize: 10, color: T.inkFaint }}>Intensidade (dB)</th><th style={{ fontSize: 10, color: T.inkFaint }}>% Monossílabas</th><th style={{ fontSize: 10, color: T.inkFaint }}>% Dissílabas</th><th style={{ fontSize: 10, color: T.inkFaint }}>Masc. (dB SN)</th>
        </tr></thead><tbody>
          <tr><td style={{ fontSize: 11, color: "#C0392B" }}>O.D.</td><td><F v={value.iprf.od.intensidade} path="iprf.od.intensidade" /></td><td><F v={value.iprf.od.mono} path="iprf.od.mono" /></td><td><F v={value.iprf.od.dissi} path="iprf.od.dissi" /></td><td><F v={value.iprf.od.masc} path="iprf.od.masc" /></td></tr>
          <tr><td style={{ fontSize: 11, color: "#2C6E62" }}>O.E.</td><td><F v={value.iprf.oe.intensidade} path="iprf.oe.intensidade" /></td><td><F v={value.iprf.oe.mono} path="iprf.oe.mono" /></td><td><F v={value.iprf.oe.dissi} path="iprf.oe.dissi" /></td><td><F v={value.iprf.oe.masc} path="iprf.oe.masc" /></td></tr>
        </tbody></table>
      </div>
      <div className="sm:col-span-2 flex flex-col items-start gap-1">
        <div style={{ fontSize: 11, color: T.inkSoft }}>Índice Percentual de Reconhecimento de Fala — barra cheia: monossílabas · barra clara: dissílabas</div>
        <IPRFBarChart iprf={value.iprf} />
      </div>
    </div>
  );
}
function LogoaudioReadView({ value }) {
  return (
    <div className="flex flex-col gap-2">
      <ReadRow label="SRT/LRF · LDF · Masc — OD" value={`${value.od.srt || "—"} dB · ${value.od.ldf || "—"} dB · ${value.od.masc || "—"} dB SN`} />
      <ReadRow label="SRT/LRF · LDF · Masc — OE" value={`${value.oe.srt || "—"} dB · ${value.oe.ldf || "—"} dB · ${value.oe.masc || "—"} dB SN`} />
      <ReadRow label="IPRF — OD" value={`${value.iprf.od.intensidade || "—"} dB · Mono ${value.iprf.od.mono || "—"}% · Dissi ${value.iprf.od.dissi || "—"}% · Masc ${value.iprf.od.masc || "—"} dB SN`} />
      <ReadRow label="IPRF — OE" value={`${value.iprf.oe.intensidade || "—"} dB · Mono ${value.iprf.oe.mono || "—"}% · Dissi ${value.iprf.oe.dissi || "—"}% · Masc ${value.iprf.oe.masc || "—"} dB SN`} />
      <IPRFBarChart iprf={value.iprf} />
    </div>
  );
}

/* Timpanograma — curva estimada a partir da pressão de pico, complacência e
   largura registradas (aproximação gaussiana; tipo B é exibido como curva
   plana, conforme classificação informada). */
function TimpanogramChart({ pressao, complacencia, largura, classificacao, color, label }) {
  const W = 260, H = 150, padL = 26, padR = 8, padT = 10, padB = 22;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const xMin = -300, xMax = 200, yMax = 2.0;
  const xFor = (p) => padL + ((p - xMin) / (xMax - xMin)) * plotW;
  const yFor = (c) => padT + plotH - (Math.min(yMax, Math.max(0, c)) / yMax) * plotH;
  const p = parseFloat(pressao), c = parseFloat(complacencia), w = parseFloat(largura);
  const isFlat = /b/i.test(classificacao || "");
  const hasData = !isNaN(c) && c > 0;
  const sigma = (!isNaN(w) && w > 0 ? w : 100) / 2.355;
  const peakX = !isNaN(p) ? p : 0;
  const pts = [];
  if (hasData && !isFlat) {
    for (let x = xMin; x <= xMax; x += 10) {
      const y = 0.05 + c * Math.exp(-((x - peakX) ** 2) / (2 * sigma * sigma));
      pts.push([xFor(x), yFor(y)]);
    }
  } else if (hasData && isFlat) {
    for (let x = xMin; x <= xMax; x += 10) pts.push([xFor(x), yFor(0.15)]);
  }
  return (
    <div className="flex flex-col items-center gap-1">
      <div style={{ fontSize: 11, fontWeight: 600, color }}>{label}</div>
      <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: 260, background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 6 }}>
        {[0, 0.5, 1.0, 1.5, 2.0].map((v) => (
          <g key={v}>
            <line x1={padL} x2={W - padR} y1={yFor(v)} y2={yFor(v)} stroke={T.lineSoft} strokeWidth="0.6" />
            <text x={padL - 4} y={yFor(v) + 3} fontSize="7" textAnchor="end" fill={T.inkFaint}>{v}</text>
          </g>
        ))}
        {[-300, -200, -100, 0, 100, 200].map((v) => (
          <g key={v}>
            <line x1={xFor(v)} x2={xFor(v)} y1={padT} y2={H - padB} stroke={T.lineSoft} strokeWidth="0.6" />
            <text x={xFor(v)} y={H - padB + 12} fontSize="7" textAnchor="middle" fill={T.inkFaint}>{v}</text>
          </g>
        ))}
        {pts.length > 1 && <polyline points={pts.map((pt) => pt.join(",")).join(" ")} fill="none" stroke={color} strokeWidth="1.5" />}
        {!hasData && <text x={W / 2} y={H / 2} fontSize="8.5" textAnchor="middle" fill={T.inkFaint}>Sem dados suficientes</text>}
      </svg>
    </div>
  );
}

function TimpanometriaEditor({ value, onChange, disabled }) {
  const rows = [["pressao","Pressão (daPa)"],["complacencia","Complacência vol. OM (mL/mmho)"],["volOE","Vol. OE (mL/mmho)"],["gradiente","Gradiente timp. (mL/mmho)"],["largura","Largura timp. (daPa)"]];
  return (
    <div className="flex flex-col gap-3">
      <table style={{ borderCollapse: "collapse", width: "100%" }}>
        <thead><tr><th></th><th style={{ fontSize: 10, color: "#C0392B" }}>Orelha direita</th><th style={{ fontSize: 10, color: "#2C6E62" }}>Orelha esquerda</th></tr></thead>
        <tbody>
          {rows.map(([k, label]) => (
            <tr key={k}>
              <td style={{ fontSize: 11, color: T.inkSoft, padding: "2px 4px" }}>{label}</td>
              <td style={{ padding: 1 }}><Cell value={value.od[k]} disabled={disabled} onChange={(v) => onChange(`od.${k}`, v)} /></td>
              <td style={{ padding: 1 }}><Cell value={value.oe[k]} disabled={disabled} onChange={(v) => onChange(`oe.${k}`, v)} /></td>
            </tr>
          ))}
        </tbody>
      </table>
      <div className="grid sm:grid-cols-2 gap-4">
        <TimpanogramChart pressao={value.od.pressao} complacencia={value.od.complacencia} largura={value.od.largura} classificacao={value.classificacao.od} color={OD_COLOR} label="Orelha direita" />
        <TimpanogramChart pressao={value.oe.pressao} complacencia={value.oe.complacencia} largura={value.oe.largura} classificacao={value.classificacao.oe} color={OE_COLOR} label="Orelha esquerda" />
      </div>
      <SubBlock title="Classificação do timpanograma">
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="OD"><TextInput value={value.classificacao.od} disabled={disabled} onChange={(v) => onChange("classificacao.od", v)} /></Field>
          <Field label="OE"><TextInput value={value.classificacao.oe} disabled={disabled} onChange={(v) => onChange("classificacao.oe", v)} /></Field>
        </div>
        <Field label="OBS"><TextArea rows={2} value={value.classificacao.obs} disabled={disabled} onChange={(v) => onChange("classificacao.obs", v)} /></Field>
        <Field label="Referência"><TextInput value={value.classificacao.referencia} disabled={disabled} onChange={(v) => onChange("classificacao.referencia", v)} /></Field>
      </SubBlock>
    </div>
  );
}
function TimpanometriaReadView({ value }) {
  const rows = [["pressao","Pressão"],["complacencia","Complacência vol. OM"],["volOE","Vol. OE"],["gradiente","Gradiente timp."],["largura","Largura timp."]];
  return (
    <div className="flex flex-col gap-2">
      {rows.map(([k, label]) => <ReadRow key={k} label={label} value={`OD: ${value.od[k] || "—"} · OE: ${value.oe[k] || "—"}`} />)}
      <div className="grid sm:grid-cols-2 gap-4 my-1">
        <TimpanogramChart pressao={value.od.pressao} complacencia={value.od.complacencia} largura={value.od.largura} classificacao={value.classificacao.od} color={OD_COLOR} label="Orelha direita" />
        <TimpanogramChart pressao={value.oe.pressao} complacencia={value.oe.complacencia} largura={value.oe.largura} classificacao={value.classificacao.oe} color={OE_COLOR} label="Orelha esquerda" />
      </div>
      <ReadRow label="Classificação do timpanograma" value={`OD: ${value.classificacao.od || "—"} · OE: ${value.classificacao.oe || "—"}`} />
      <ReadRow label="Obs. / Referência" value={[value.classificacao.obs, value.classificacao.referencia].filter(Boolean).join("  —  ")} />
    </div>
  );
}

/* Reflex threshold line chart — RACL/RAIL por frequência, ambas as orelhas */
function ReflexLineChart({ value }) {
  const W = 300, H = 160, padL = 30, padR = 8, padT = 10, padB = 22;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const freqs = REFLEX_FREQS.map(Number);
  const xFor = (f) => padL + (Math.log2(f / 500) / Math.log2(4000 / 500)) * plotW;
  const yFor = (db) => padT + (Math.min(120, Math.max(0, db)) / 120) * plotH;
  const series = [
    { key: "racl", ear: "od", color: OD_COLOR, dash: "0", label: "OD RACL" },
    { key: "rail", ear: "od", color: OD_COLOR, dash: "4,3", label: "OD RAIL" },
    { key: "racl", ear: "oe", color: OE_COLOR, dash: "0", label: "OE RACL" },
    { key: "rail", ear: "oe", color: OE_COLOR, dash: "4,3", label: "OE RAIL" },
  ];
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: 300, background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 6 }}>
      {[0, 30, 60, 90, 120].map((v) => (
        <g key={v}>
          <line x1={padL} x2={W - padR} y1={yFor(v)} y2={yFor(v)} stroke={T.lineSoft} strokeWidth="0.6" />
          <text x={padL - 4} y={yFor(v) + 3} fontSize="7" textAnchor="end" fill={T.inkFaint}>{v}</text>
        </g>
      ))}
      {freqs.map((f) => <text key={f} x={xFor(f)} y={H - padB + 12} fontSize="7.5" textAnchor="middle" fill={T.inkFaint}>{f}</text>)}
      {series.map((s) => {
        const pts = REFLEX_FREQS.filter((f) => value[s.ear][f][s.key] !== "" && !isNaN(value[s.ear][f][s.key])).map((f) => [xFor(Number(f)), yFor(Number(value[s.ear][f][s.key]))]);
        return pts.length > 0 ? (
          <g key={s.ear + s.key}>
            {pts.length > 1 && <polyline points={pts.map((p) => p.join(",")).join(" ")} fill="none" stroke={s.color} strokeWidth="1.4" strokeDasharray={s.dash} />}
            {pts.map(([x, y], i) => <circle key={i} cx={x} cy={y} r="2.6" fill={s.color} />)}
          </g>
        ) : null;
      })}
    </svg>
  );
}

function ReflexosEditor({ value, onChange, disabled }) {
  const Table = ({ ear, earLabel, color }) => (
    <div>
      <div style={{ fontSize: 11.5, fontWeight: 600, color, marginBottom: 4 }}>{earLabel}</div>
      <table style={{ borderCollapse: "collapse", width: "100%" }}>
        <thead><tr>
          <th style={{ fontSize: 9.5, color: T.inkFaint }}>Freq.</th>
          <th style={{ fontSize: 9.5, color: T.inkFaint }}>Limiar psicoac.</th>
          <th style={{ fontSize: 9.5, color: T.inkFaint }}>RACL</th>
          <th style={{ fontSize: 9.5, color: T.inkFaint }}>Dif.</th>
          <th style={{ fontSize: 9.5, color: T.inkFaint }}>RAIL</th>
        </tr></thead>
        <tbody>
          {REFLEX_FREQS.map((f) => (
            <tr key={f}>
              <td style={{ fontSize: 10, color: T.inkSoft, textAlign: "center" }}>{f === "500" ? ".5k" : Number(f) / 1000 + "k"} Hz</td>
              {["limiar","racl","dif","rail"].map((col) => (
                <td key={col} style={{ padding: 1 }}><Cell value={value[ear][f][col]} disabled={disabled} onChange={(v) => onChange(`${ear}.${f}.${col}`, v)} /></td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
  return (
    <div className="flex flex-col gap-4">
      <div className="grid sm:grid-cols-2 gap-4">
        <Table ear="od" earLabel="Orelha direita" color="#C0392B" />
        <Table ear="oe" earLabel="Orelha esquerda" color="#2C6E62" />
      </div>
      <div className="flex flex-col items-start gap-1">
        <div style={{ fontSize: 11, color: T.inkSoft }}>Linha cheia: RACL · Linha tracejada: RAIL — vermelho OD, azul OE</div>
        <ReflexLineChart value={value} />
      </div>
      <SubBlock title="Classificação do Reflexo Acústico Contralateral">
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="OD"><TextInput value={value.classificacao.od} disabled={disabled} onChange={(v) => onChange("classificacao.od", v)} /></Field>
          <Field label="OE"><TextInput value={value.classificacao.oe} disabled={disabled} onChange={(v) => onChange("classificacao.oe", v)} /></Field>
        </div>
        <Field label="OBS"><TextArea rows={2} value={value.classificacao.obs} disabled={disabled} onChange={(v) => onChange("classificacao.obs", v)} /></Field>
        <Field label="Referência"><TextInput value={value.classificacao.referencia} disabled={disabled} onChange={(v) => onChange("classificacao.referencia", v)} /></Field>
      </SubBlock>
    </div>
  );
}
function ReflexosReadView({ value }) {
  const line = (ear) => REFLEX_FREQS.map((f) => {
    const c = value[ear][f];
    return c.limiar || c.racl || c.dif || c.rail ? `${f}Hz(limiar ${c.limiar || "—"}, RACL ${c.racl || "—"}, dif ${c.dif || "—"}, RAIL ${c.rail || "—"})` : null;
  }).filter(Boolean).join(" · ");
  return (
    <div className="flex flex-col gap-2">
      <ReadRow label="Orelha direita" value={line("od")} />
      <ReadRow label="Orelha esquerda" value={line("oe")} />
      <ReflexLineChart value={value} />
      <ReadRow label="Classificação do Reflexo Acústico Contralateral" value={`OD: ${value.classificacao.od || "—"} · OE: ${value.classificacao.oe || "—"}`} />
      <ReadRow label="Obs. / Referência" value={[value.classificacao.obs, value.classificacao.referencia].filter(Boolean).join("  —  ")} />
    </div>
  );
}

function audiogramTextLines(value) {
  const fmt = (ear, cond) => FREQS_AUD.filter((f) => value[ear][cond][f]).map((f) => `${f}Hz:${value[ear][cond][f]}dB`).join(" ");
  const pta = (ear) => { const p = computePTA(value[ear].va); return p != null ? `${p} dB (${classifyPTA(p)})` : "—"; };
  const lossType = (ear) => classifyLossType(computePTA(value[ear].va), computePTA(value[ear].vo)) || "—";
  const advisories = computeMaskingAdvice(value.od, value.oe);
  const maskingLine = advisories.length
    ? advisories.map((a) => `${a.ear} ${a.freq}Hz ${a.cond === "va" ? "VA" : "VO"}: ${a.suggestion}`).join(" · ")
    : "Nenhuma indicação de mascaramento pelas regras de atenuação interaural (estimativa)";
  return [
    ["OD — via aérea", fmt("od", "va") || "—"], ["OD — via óssea", fmt("od", "vo") || "—"],
    ["OE — via aérea", fmt("oe", "va") || "—"], ["OE — via óssea", fmt("oe", "vo") || "—"],
    ["OD — PTA (500–4k)", pta("od")], ["OE — PTA (500–4k)", pta("oe")],
    ["OD — Tipo de perda (estimado)", lossType("od")], ["OE — Tipo de perda (estimado)", lossType("oe")],
    ["Mascaramento (estimativa)", maskingLine],
    ["Mascaramento aplicado (registrado)", maskingLogText(value.mascaramentos)],
    ["Laudo OD", value.laudo.od || "—"], ["Laudo OE", value.laudo.oe || "—"],
    ["Laudo OBS", value.laudo.obs || "—"], ["Referência", value.laudo.referencia || "—"],
  ];
}
function logoaudioTextLines(value) {
  return [
    ["SRT/LRF · LDF · Masc — OD", `${value.od.srt || "—"} dB · ${value.od.ldf || "—"} dB · ${value.od.masc || "—"} dB SN`],
    ["SRT/LRF · LDF · Masc — OE", `${value.oe.srt || "—"} dB · ${value.oe.ldf || "—"} dB · ${value.oe.masc || "—"} dB SN`],
    ["IPRF — OD", `${value.iprf.od.intensidade || "—"} dB · Mono ${value.iprf.od.mono || "—"}% · Dissi ${value.iprf.od.dissi || "—"}%`],
    ["IPRF — OE", `${value.iprf.oe.intensidade || "—"} dB · Mono ${value.iprf.oe.mono || "—"}% · Dissi ${value.iprf.oe.dissi || "—"}%`],
  ];
}
function timpanometriaTextLines(value) {
  const rows = [["pressao","Pressão"],["complacencia","Complacência vol. OM"],["volOE","Vol. OE"],["gradiente","Gradiente timp."],["largura","Largura timp."]];
  return [
    ...rows.map(([k, label]) => [label, `OD: ${value.od[k] || "—"} · OE: ${value.oe[k] || "—"}`]),
    ["Classificação do timpanograma", `OD: ${value.classificacao.od || "—"} · OE: ${value.classificacao.oe || "—"}`],
    ["Obs. / Referência", [value.classificacao.obs, value.classificacao.referencia].filter(Boolean).join(" — ") || "—"],
  ];
}
function reflexosTextLines(value) {
  const line = (ear) => REFLEX_FREQS.map((f) => {
    const c = value[ear][f];
    return c.limiar || c.racl || c.dif || c.rail ? `${f}Hz(limiar ${c.limiar || "—"},RACL ${c.racl || "—"},dif ${c.dif || "—"},RAIL ${c.rail || "—"})` : null;
  }).filter(Boolean).join(" ");
  return [
    ["Orelha direita", line("od") || "—"], ["Orelha esquerda", line("oe") || "—"],
    ["Classificação Reflexo Contralateral", `OD: ${value.classificacao.od || "—"} · OE: ${value.classificacao.oe || "—"}`],
    ["Obs. / Referência", [value.classificacao.obs, value.classificacao.referencia].filter(Boolean).join(" — ") || "—"],
  ];
}

/* ---------------------------------------------------------------------
   Generic schema-driven field renderer (edit mode)
------------------------------------------------------------------------ */
function SchemaField({ node, data, onChange, disabled }) {
  if (node.type === "group") {
    return (
      <SubBlock title={node.title}>
        {node.children.map((c, i) => <SchemaField key={i} node={c} data={data} onChange={onChange} disabled={disabled} />)}
      </SubBlock>
    );
  }
  if (node.type === "custom") {
    const Comp = node.Edit;
    return <Comp value={_.get(data, node.path)} disabled={disabled}
      onChange={(subpath, v) => onChange(subpath ? `${node.path}.${subpath}` : node.path, v)} />;
  }
  if (node.type === "checks") {
    return (
      <div className="flex flex-col gap-2">
        <CheckRow options={node.options} value={_.get(data, node.path)} disabled={disabled}
          onChange={(k, v) => onChange(`${node.path}.${k}`, v)} />
        {node.extra && (
          <Field label={node.extra.label}>
            <TextInput value={_.get(data, node.extra.path)} disabled={disabled} onChange={(v) => onChange(node.extra.path, v)} />
          </Field>
        )}
      </div>
    );
  }
  if (node.type === "textarea") {
    return (
      <Field label={node.label}>
        <TextArea rows={node.rows} value={_.get(data, node.path)} disabled={disabled} onChange={(v) => onChange(node.path, v)} />
      </Field>
    );
  }
  return (
    <Field label={node.label}>
      <TextInput value={_.get(data, node.path)} disabled={disabled} onChange={(v) => onChange(node.path, v)} />
    </Field>
  );
}

function ClinicoForm({ sections, clinico, setClinico, startNumber = 3 }) {
  const [open, setOpen] = useState({ 0: true });
  const toggle = (i) => setOpen((s) => ({ ...s, [i]: !s[i] }));
  const onChange = (path, value) => setClinico((prev) => _.set(_.cloneDeep(prev), path, value));
  const allOpen = Object.values(open).filter(Boolean).length >= sections.length;
  return (
    <div className="flex flex-col gap-3">
      <div className="flex justify-end">
        <button
          onClick={() => setOpen(allOpen ? {} : Object.fromEntries(sections.map((_s, i) => [i, true])))}
          style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontSize: 12.5, color: T.accent, fontWeight: 500 }}
        >
          {allOpen ? "Recolher tudo" : "Expandir tudo"}
        </button>
      </div>
      {sections.map((s, i) => (
        <Section key={i} number={i + startNumber} title={s.title} open={!!open[i]} onToggle={() => toggle(i)}>
          {s.children.map((c, j) => <SchemaField key={j} node={c} data={clinico} onChange={onChange} />)}
        </Section>
      ))}
    </div>
  );
}

/* ---------------------------------------------------------------------
   Generic schema-driven read-only rendering (view + print)
------------------------------------------------------------------------ */
function readNodeString(node, data) {
  if (node.type === "group") {
    return node.children.map((c) => readNodeString(c, data)).filter(Boolean).join("  —  ");
  }
  if (node.type === "checks") {
    const marks = node.options.filter(([k]) => _.get(data, `${node.path}.${k}`)).map(([, l]) => l).join(" · ");
    const extra = node.extra ? _.get(data, node.extra.path) : "";
    return [marks, extra].filter(Boolean).join("  —  ");
  }
  return _.get(data, node.path) || "";
}
function ReadRow({ label, value }) {
  return (
    <div className="flex flex-col gap-0.5">
      {label && <span style={{ fontSize: 11.5, color: T.inkFaint, fontFamily: "'IBM Plex Sans', sans-serif" }}>{label}</span>}
      <span style={{ fontSize: 14, color: T.ink, fontFamily: "'IBM Plex Sans', sans-serif" }}>
        {value || <span style={{ color: T.inkFaint, fontStyle: "italic" }}>—</span>}
      </span>
    </div>
  );
}
function PrintSection({ number, title, children }) {
  return (
    <div style={{ marginBottom: 14, breakInside: "avoid" }}>
      <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontWeight: 600, fontSize: 13, color: T.accentDeep, background: T.accentSoft, padding: "4px 8px", marginBottom: 8 }}>
        {number}. {title}
      </div>
      {children}
    </div>
  );
}
function ClinicoReadView({ sections, clinico, startNumber = 3 }) {
  return (
    <>
      {sections.map((s, i) => (
        <PrintSection key={i} number={i + startNumber} title={s.title}>
          <div className="flex flex-col gap-2">
            {s.children.map((c, j) => {
              if (c.type === "custom") {
                const Comp = c.Read;
                return <Comp key={j} value={_.get(clinico, c.path)} />;
              }
              return c.type === "checks" && !c.label ? (
                <ReadRow key={j} value={readNodeString(c, clinico)} />
              ) : (
                <ReadRow key={j} label={c.title || c.label} value={readNodeString(c, clinico)} />
              );
            })}
          </div>
        </PrintSection>
      ))}
    </>
  );
}

/* ---------------------------------------------------------------------
   Direct PDF export (no print dialog) — builds a real, selectable-text
   PDF via jsPDF, loaded on demand from a CDN. Falls back to the
   browser's print-to-PDF if the library can't be loaded.
------------------------------------------------------------------------ */
function sectionTextLines(section, clinico) {
  const lines = [];
  section.children.forEach((c) => {
    if (c.type === "custom") {
      lines.push(...c.textLines(_.get(clinico, c.path)));
    } else if (c.type === "checks" && !c.label) {
      lines.push([null, readNodeString(c, clinico)]);
    } else {
      lines.push([c.title || c.label || null, readNodeString(c, clinico)]);
    }
  });
  return lines;
}

async function generateRecordPDF(record) {
  const { jsPDF } = await import("https://cdn.jsdelivr.net/npm/jspdf@2.5.2/+esm");
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const pageW = 210, pageH = 297, mL = 16, mR = 16, mT = 18, mB = 18;
  const contentW = pageW - mL - mR;
  const pos = { y: mT };
  const accent = [30, 78, 69], accentSoft = [220, 234, 229], ink = [29, 42, 40], inkSoft = [84, 101, 95];

  const ensure = (need) => { if (pos.y + need > pageH - mB) { doc.addPage(); pos.y = mT; } };
  const sectionHeader = (num, title) => {
    ensure(13);
    doc.setFillColor(...accentSoft);
    doc.rect(mL, pos.y - 4.2, contentW, 7, "F");
    doc.setFont(undefined, "bold"); doc.setFontSize(10.5); doc.setTextColor(...accent);
    doc.text(`${num}. ${title}`, mL + 2, pos.y);
    pos.y += 9;
  };
  const kv = (label, value) => {
    const text = (value ?? "").toString().trim() || "—";
    if (label) {
      ensure(4);
      doc.setFont(undefined, "bold"); doc.setFontSize(8); doc.setTextColor(...inkSoft);
      doc.text(label, mL, pos.y); pos.y += 3.8;
    }
    doc.setFont(undefined, "normal"); doc.setFontSize(9.6); doc.setTextColor(...ink);
    doc.splitTextToSize(text, contentW).forEach((line) => { ensure(4.4); doc.text(line, mL, pos.y); pos.y += 4.4; });
    pos.y += 2;
  };
  const kvGrid = (pairs) => {
    const colW = (contentW - 6) / 2;
    for (let i = 0; i < pairs.length; i += 2) {
      const a = pairs[i], b = pairs[i + 1];
      ensure(9);
      doc.setFont(undefined, "bold"); doc.setFontSize(8); doc.setTextColor(...inkSoft);
      if (a && a[0]) doc.text(a[0], mL, pos.y);
      if (b && b[0]) doc.text(b[0], mL + colW + 6, pos.y);
      pos.y += 3.8;
      doc.setFont(undefined, "normal"); doc.setFontSize(9.6); doc.setTextColor(...ink);
      if (a) doc.text((a[1] || "—").toString(), mL, pos.y);
      if (b) doc.text((b[1] || "—").toString(), mL + colW + 6, pos.y);
      pos.y += 6;
    }
  };

  const id = record.identificacao, p = record.paciente;
  doc.setFont(undefined, "bold"); doc.setFontSize(13); doc.setTextColor(...ink);
  doc.text(id.instituicao || "", pageW / 2, pos.y, { align: "center" }); pos.y += 6;
  doc.setFont(undefined, "normal"); doc.setFontSize(9.5); doc.setTextColor(...inkSoft);
  doc.text(`${id.curso || ""} — ${id.disciplina || ""}`, pageW / 2, pos.y, { align: "center" }); pos.y += 6;
  doc.setFont(undefined, "bold"); doc.setFontSize(12); doc.setTextColor(...accent);
  doc.text(`Protocolo Padronizado de Registro de Atendimento — ${AREAS[record.area]?.label || ""}`, pageW / 2, pos.y, { align: "center", maxWidth: contentW }); pos.y += 10;

  sectionHeader(1, "Identificação");
  kvGrid([
    ["Fonoaudiólogo(a) responsável", id.supervisor], ["CRFa", id.crfa],
    ["Estagiário(a) (se houver)", id.estagiario], ["Matrícula", id.matricula],
    ["Data", id.data], ["Horário", id.horario],
    ["Sessão nº", id.sessaoNum], ["Área", AREAS[record.area]?.label],
  ]);
  pos.y += 1;

  sectionHeader(2, "Paciente");
  kvGrid([
    ["Nome", p.nome], ["Nascimento", p.nascimento],
    ["Idade", p.idade], ["Responsável", p.responsavel],
    ["Hipótese diagnóstica", p.hipotese], ["CID", p.cid],
    ["Tempo de acompanhamento", p.tempoAcompanhamento], [null, null],
  ]);
  pos.y += 1;

  const area = AREAS[record.area];
  area.sections.forEach((s, i) => {
    sectionHeader(i + 3, s.title);
    sectionTextLines(s, record.clinico).forEach(([label, value]) => kv(label, value));
  });

  const n = area.sections.length;
  sectionHeader(n + 3, "Procedimentos Realizados"); kv(null, record.procedimentos);
  sectionHeader(n + 4, "Evolução Clínica"); kv(null, record.evolucao);
  sectionHeader(n + 5, "Orientações à Família"); kv(null, record.orientacoes);
  sectionHeader(n + 6, "Plano para Próxima Sessão"); kv(null, record.plano);
  sectionHeader(n + 7, "Raciocínio Clínico");
  kvGrid([
    ["Principal alteração", record.raciocinio.alteracao], ["Evidências", record.raciocinio.evidencias],
    ["Objetivo alcançado", record.raciocinio.objetivoAlcancado], ["Conduta seguinte", record.raciocinio.conduta],
  ]);

  sectionHeader(n + 8, "Estudo de Caso (desta sessão)");
  kv("Discutido em estudo de caso?", record.estudoDeCasoSessao?.discutido ? "Sim" : "Não");
  if (record.estudoDeCasoSessao?.resumo) kv("Resumo da discussão", record.estudoDeCasoSessao.resumo);

  sectionHeader(n + 9, "Validação da Supervisão");
  kv("Status", record.validacaoSupervisao?.status === "validado" ? "Validado" : "Pendente de validação");
  kvGrid([
    ["Supervisor(a)", record.validacaoSupervisao?.supervisorNome], ["CRFa", record.validacaoSupervisao?.crfaSupervisor],
    ["Data da validação", record.validacaoSupervisao?.dataValidacao], [null, null],
  ]);
  if (record.validacaoSupervisao?.comentario) kv("Comentário", record.validacaoSupervisao.comentario);

  sectionHeader(n + 10, "Assinaturas");
  ensure(42);
  pos.y += 18;
  const colW = (contentW - 10) / 2;
  doc.setDrawColor(...ink); doc.setLineWidth(0.3);
  doc.line(mL, pos.y, mL + colW, pos.y);
  doc.line(mL + colW + 10, pos.y, mL + colW + 10 + colW, pos.y);
  pos.y += 4;
  doc.setFont(undefined, "normal"); doc.setFontSize(9); doc.setTextColor(...ink);
  doc.text(id.supervisor || "Fonoaudiólogo(a) responsável", mL, pos.y);
  doc.text(id.estagiario || "Estagiário(a) (se houver)", mL + colW + 10, pos.y);
  pos.y += 4;
  doc.setFontSize(8); doc.setTextColor(...inkSoft);
  doc.text(`Fonoaudiólogo(a) responsável${id.crfa ? " · CRFa " + id.crfa : ""}`, mL, pos.y);
  doc.text(`Estagiário(a)${id.matricula ? " · Mat. " + id.matricula : ""}`, mL + colW + 10, pos.y);
  pos.y += 8;
  doc.text("Data: ____/____/________", mL, pos.y);
  doc.text("Data: ____/____/________", mL + colW + 10, pos.y);

  const filename = `atendimento_${(p.nome || "paciente").replace(/[^a-zA-Z0-9]+/g, "_")}_${id.data || "sem_data"}.pdf`;
  doc.save(filename);
}

const AUDIOLOGY_AREAS = ["audiologia_adulto", "audiologia_infantil"];
function isAudiologyRecord(record) {
  return AUDIOLOGY_AREAS.includes(record?.area) && record?.clinico?.audiograma;
}

/* Compact, standalone audiological report (laudo) — a short document
   separate from the full session record, ready to hand to the patient or
   referring physician and to be signed by the supervising audiologist. */
async function generateLaudoPDF(record) {
  const { jsPDF } = await import("https://cdn.jsdelivr.net/npm/jspdf@2.5.2/+esm");
  const doc = new jsPDF({ unit: "mm", format: "a4" });
  const pageW = 210, mL = 20, mR = 20, mT = 20;
  const contentW = pageW - mL - mR;
  const pos = { y: mT };
  const ink = [29, 42, 40], inkSoft = [84, 101, 95], accent = [30, 78, 69];

  const id = record.identificacao, p = record.paciente;
  const a = record.clinico.audiograma;
  const l = record.clinico.logoaudio, t = record.clinico.timpanometria;

  doc.setFont(undefined, "bold"); doc.setFontSize(13); doc.setTextColor(...ink);
  doc.text(id.instituicao || "", pageW / 2, pos.y, { align: "center" }); pos.y += 6;
  doc.setFont(undefined, "normal"); doc.setFontSize(9.5); doc.setTextColor(...inkSoft);
  doc.text(`${id.curso || ""} — ${id.disciplina || ""}`, pageW / 2, pos.y, { align: "center" }); pos.y += 10;
  doc.setFont(undefined, "bold"); doc.setFontSize(14); doc.setTextColor(...accent);
  doc.text("LAUDO AUDIOLÓGICO", pageW / 2, pos.y, { align: "center" }); pos.y += 12;

  doc.setDrawColor(210, 219, 216); doc.line(mL, pos.y, pageW - mR, pos.y); pos.y += 8;

  const field = (label, value) => {
    doc.setFont(undefined, "bold"); doc.setFontSize(9); doc.setTextColor(...inkSoft);
    doc.text(label, mL, pos.y);
    doc.setFont(undefined, "normal"); doc.setFontSize(10); doc.setTextColor(...ink);
    doc.text((value || "—").toString(), mL + 42, pos.y);
    pos.y += 6.5;
  };
  field("Paciente:", p.nome);
  field("Data de nascimento:", p.nascimento);
  field("Idade:", p.idade);
  field("Data do exame:", id.data);
  pos.y += 4;

  const examsDone = [];
  if (a) examsDone.push("Audiometria tonal liminar");
  if (l && (l.od.srt || l.iprf.od.mono)) examsDone.push("Logoaudiometria");
  if (t && (t.od.pressao || t.classificacao.od)) examsDone.push("Imitanciometria");
  if (record.clinico.reflexos?.classificacao?.od || record.clinico.reflexos?.classificacao?.oe) examsDone.push("Pesquisa de reflexos acústicos");
  field("Exames realizados:", examsDone.join(", ") || "—");
  pos.y += 4;

  doc.setFont(undefined, "bold"); doc.setFontSize(10.5); doc.setTextColor(...accent);
  doc.text("Achados e conclusão", mL, pos.y); pos.y += 6;
  doc.setFont(undefined, "normal"); doc.setFontSize(10); doc.setTextColor(...ink);
  const laudoText = (a?.laudo?.obs && a.laudo.obs.trim()) || generateLaudoText(a, l, t);
  doc.splitTextToSize(laudoText, contentW).forEach((line) => { doc.text(line, mL, pos.y); pos.y += 5.2; });
  pos.y += 4;

  if (a?.laudo?.od || a?.laudo?.oe) {
    if (a.laudo.od) { doc.setFont(undefined, "bold"); doc.setFontSize(9); doc.text("OD:", mL, pos.y); doc.setFont(undefined, "normal"); doc.splitTextToSize(a.laudo.od, contentW - 12).forEach((line, i) => doc.text(line, mL + 12, pos.y + i * 5)); pos.y += 6; }
    if (a.laudo.oe) { doc.setFont(undefined, "bold"); doc.setFontSize(9); doc.text("OE:", mL, pos.y); doc.setFont(undefined, "normal"); doc.splitTextToSize(a.laudo.oe, contentW - 12).forEach((line, i) => doc.text(line, mL + 12, pos.y + i * 5)); pos.y += 6; }
  }
  if (a?.laudo?.referencia) { doc.setFontSize(8.5); doc.setTextColor(...inkSoft); doc.text(`Referência: ${a.laudo.referencia}`, mL, pos.y); pos.y += 8; }

  pos.y += 20;
  const colW = (contentW - 10) / 2;
  doc.setDrawColor(...ink); doc.setLineWidth(0.3);
  doc.line(mL, pos.y, mL + colW, pos.y);
  doc.line(mL + colW + 10, pos.y, mL + colW + 10 + colW, pos.y);
  pos.y += 4;
  doc.setFont(undefined, "normal"); doc.setFontSize(9); doc.setTextColor(...ink);
  doc.text(id.supervisor || "Fonoaudiólogo(a) supervisor(a)", mL, pos.y);
  doc.text(id.estagiario || "Estagiário(a) responsável", mL + colW + 10, pos.y);
  pos.y += 4;
  doc.setFontSize(8); doc.setTextColor(...inkSoft);
  doc.text(id.crfa ? `CRFa ${id.crfa}` : "CRFa", mL, pos.y);
  doc.text(id.matricula ? `Mat. ${id.matricula}` : "", mL + colW + 10, pos.y);

  doc.save(`laudo_audiologico_${(p.nome || "paciente").replace(/[^a-zA-Z0-9]+/g, "_")}_${id.data || "sem_data"}.pdf`);
}

function RecordReadView({ record }) {
  const id = record.identificacao, p = record.paciente;
  return (
    <div style={{ fontFamily: "'IBM Plex Sans', sans-serif" }}>
      <PrintSection number={1} title="Identificação">
        <div className="grid sm:grid-cols-3 gap-3">
          <ReadRow label="Área" value={AREAS[record.area]?.label} />
          <ReadRow label="Instituição" value={id.instituicao} />
          <ReadRow label="Curso" value={id.curso} />
          <ReadRow label="Disciplina" value={id.disciplina} />
          <ReadRow label="Fonoaudiólogo(a) responsável" value={id.supervisor} />
          <ReadRow label="CRFa" value={id.crfa} />
          <ReadRow label="Estagiário(a) (se houver)" value={id.estagiario} />
          <ReadRow label="Matrícula" value={id.matricula} />
          <ReadRow label="Data" value={id.data} />
          <ReadRow label="Horário" value={id.horario} />
          <ReadRow label="Sessão nº" value={id.sessaoNum} />
        </div>
      </PrintSection>
      <PrintSection number={2} title="Paciente">
        <div className="grid sm:grid-cols-3 gap-3">
          <ReadRow label="Nome" value={p.nome} />
          <ReadRow label="Nascimento" value={p.nascimento} />
          <ReadRow label="Idade" value={p.idade} />
          <ReadRow label="Responsável" value={p.responsavel} />
          <ReadRow label="Hipótese diagnóstica" value={p.hipotese} />
          <ReadRow label="CID" value={p.cid} />
          <ReadRow label="Tempo de acompanhamento" value={p.tempoAcompanhamento} />
        </div>
      </PrintSection>

      <ClinicoReadView sections={AREAS[record.area].sections} clinico={record.clinico} />

      {(() => {
        const n = AREAS[record.area].sections.length;
        return (
          <>
            <PrintSection number={n + 3} title="Procedimentos Realizados"><ReadRow value={record.procedimentos} /></PrintSection>
            <PrintSection number={n + 4} title="Evolução Clínica"><ReadRow value={record.evolucao} /></PrintSection>
            <PrintSection number={n + 5} title="Orientações à Família"><ReadRow value={record.orientacoes} /></PrintSection>
            <PrintSection number={n + 6} title="Plano para Próxima Sessão"><ReadRow value={record.plano} /></PrintSection>
            <PrintSection number={n + 7} title="Raciocínio Clínico">
              <div className="grid sm:grid-cols-2 gap-3">
                <ReadRow label="Principal alteração" value={record.raciocinio.alteracao} />
                <ReadRow label="Evidências" value={record.raciocinio.evidencias} />
                <ReadRow label="Objetivo alcançado" value={record.raciocinio.objetivoAlcancado} />
                <ReadRow label="Conduta seguinte" value={record.raciocinio.conduta} />
              </div>
            </PrintSection>
            <PrintSection number={n + 8} title="Estudo de Caso (desta sessão)">
              <ReadRow label="Discutido em estudo de caso?" value={record.estudoDeCasoSessao?.discutido ? "Sim" : "Não"} />
              {record.estudoDeCasoSessao?.resumo && <ReadRow label="Resumo da discussão" value={record.estudoDeCasoSessao.resumo} />}
            </PrintSection>
            <PrintSection number={n + 9} title="Validação da Supervisão">
              <ReadRow label="Status" value={record.validacaoSupervisao?.status === "validado" ? "✓ Validado" : "Pendente de validação"} />
              <div className="grid sm:grid-cols-2 gap-3">
                <ReadRow label="Supervisor(a)" value={record.validacaoSupervisao?.supervisorNome} />
                <ReadRow label="CRFa" value={record.validacaoSupervisao?.crfaSupervisor} />
                <ReadRow label="Data da validação" value={record.validacaoSupervisao?.dataValidacao} />
              </div>
              {record.validacaoSupervisao?.comentario && <ReadRow label="Comentário" value={record.validacaoSupervisao.comentario} />}
            </PrintSection>
            <PrintSection number={n + 10} title="Assinaturas">
              <div className="grid sm:grid-cols-2 gap-8" style={{ marginTop: 30 }}>
                <div style={{ textAlign: "center" }}>
                  <div style={{ borderTop: "1px solid " + T.ink, paddingTop: 6, fontSize: 13 }}>{id.supervisor || "Fonoaudiólogo(a) responsável"}</div>
                  <div style={{ fontSize: 11, color: T.inkFaint, marginTop: 2 }}>Fonoaudiólogo(a) responsável{id.crfa ? " · CRFa " + id.crfa : ""}</div>
                  <div style={{ fontSize: 11, color: T.inkFaint, marginTop: 10 }}>Data: ____/____/________</div>
                </div>
                <div style={{ textAlign: "center" }}>
                  <div style={{ borderTop: "1px solid " + T.ink, paddingTop: 6, fontSize: 13 }}>{id.estagiario || "Estagiário(a) (se houver)"}</div>
                  <div style={{ fontSize: 11, color: T.inkFaint, marginTop: 2 }}>Estagiário(a){id.matricula ? " · Mat. " + id.matricula : ""}</div>
                  <div style={{ fontSize: 11, color: T.inkFaint, marginTop: 10 }}>Data: ____/____/________</div>
                </div>
              </div>
            </PrintSection>
          </>
        );
      })()}
    </div>
  );
}

function RecordForm({ areaId, record, setRecord, perfil }) {
  const set = (path, value) => setRecord((prev) => _.set(_.cloneDeep(prev), path, value));
  const papelEfetivo = typeof window !== "undefined" && window.authProfile ? window.authProfile.papel : perfil?.papel;
  const nomeEfetivo = typeof window !== "undefined" && window.authProfile?.nome ? window.authProfile.nome : perfil?.nome;
  const souSupervisor = papelEfetivo === "supervisor" || papelEfetivo === "admin";
  const validar = () => {
    set("validacaoSupervisao.status", "validado");
    set("validacaoSupervisao.supervisorNome", nomeEfetivo || record.validacaoSupervisao.supervisorNome);
    set("validacaoSupervisao.crfaSupervisor", perfil?.crfa || record.validacaoSupervisao.crfaSupervisor);
    set("validacaoSupervisao.dataValidacao", new Date().toISOString().slice(0, 10));
  };
  const desfazerValidacao = () => { set("validacaoSupervisao.status", "pendente"); set("validacaoSupervisao.dataValidacao", ""); };
  return (
    <div className="flex flex-col gap-3">
      <div className="flex items-center gap-3 p-3" style={{ background: T.accentSoft, borderRadius: 10 }}>
        <User size={16} color={T.accentDeep} />
        <div>
          <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontWeight: 600, fontSize: 14, color: T.ink }}>{record.paciente.nome}</div>
          <div style={{ fontSize: 11.5, color: T.inkSoft }}>{AREAS[areaId].label} · dados do paciente e anamnese ficam na ficha do paciente</div>
        </div>
      </div>

      <Section number="1" title="Identificação do Atendimento" open onToggle={() => {}}>
        <div className="grid sm:grid-cols-2 gap-3">
          <Field label="Instituição" span={2}><TextInput value={record.identificacao.instituicao} onChange={(v) => set("identificacao.instituicao", v)} /></Field>
          <Field label="Curso"><TextInput value={record.identificacao.curso} onChange={(v) => set("identificacao.curso", v)} /></Field>
          <Field label="Disciplina"><TextInput value={record.identificacao.disciplina} onChange={(v) => set("identificacao.disciplina", v)} /></Field>
          <Field label="Fonoaudiólogo(a) responsável"><TextInput value={record.identificacao.supervisor} onChange={(v) => set("identificacao.supervisor", v)} /></Field>
          <Field label="CRFa"><TextInput value={record.identificacao.crfa} onChange={(v) => set("identificacao.crfa", v)} /></Field>
          <Field label="Estagiário(a) (se houver)"><TextInput value={record.identificacao.estagiario} onChange={(v) => set("identificacao.estagiario", v)} /></Field>
          <Field label="Matrícula"><TextInput value={record.identificacao.matricula} onChange={(v) => set("identificacao.matricula", v)} /></Field>
          <Field label="Data"><TextInput type="date" value={record.identificacao.data} onChange={(v) => set("identificacao.data", v)} /></Field>
          <Field label="Horário"><TextInput type="time" value={record.identificacao.horario} onChange={(v) => set("identificacao.horario", v)} /></Field>
          <Field label="Sessão nº"><TextInput type="number" value={record.identificacao.sessaoNum} onChange={(v) => set("identificacao.sessaoNum", v)} /></Field>
        </div>
      </Section>

      <ClinicoForm sections={AREAS[areaId].sections} startNumber={2} clinico={record.clinico} setClinico={(updater) => setRecord((prev) => ({ ...prev, clinico: typeof updater === "function" ? updater(prev.clinico) : updater }))} />

      {(() => {
        const n = AREAS[areaId].sections.length;
        return (
          <>
            <Section number={n + 2} title="Procedimentos Realizados" open onToggle={() => {}}>
              <TextArea rows={4} value={record.procedimentos} onChange={(v) => set("procedimentos", v)} />
            </Section>
            <Section number={n + 3} title="Evolução Clínica" open onToggle={() => {}}>
              <TextArea rows={4} value={record.evolucao} onChange={(v) => set("evolucao", v)} />
            </Section>
            <Section number={n + 4} title="Orientações à Família" open onToggle={() => {}}>
              <TextArea rows={3} value={record.orientacoes} onChange={(v) => set("orientacoes", v)} />
            </Section>
            <Section number={n + 5} title="Plano para Próxima Sessão" open onToggle={() => {}}>
              <TextArea rows={3} value={record.plano} onChange={(v) => set("plano", v)} />
            </Section>
            <Section number={n + 6} title="Raciocínio Clínico" open onToggle={() => {}}>
              <div className="grid sm:grid-cols-2 gap-3">
                <Field label="Principal alteração"><TextArea rows={2} value={record.raciocinio.alteracao} onChange={(v) => set("raciocinio.alteracao", v)} /></Field>
                <Field label="Evidências"><TextArea rows={2} value={record.raciocinio.evidencias} onChange={(v) => set("raciocinio.evidencias", v)} /></Field>
                <Field label="Objetivo alcançado"><TextArea rows={2} value={record.raciocinio.objetivoAlcancado} onChange={(v) => set("raciocinio.objetivoAlcancado", v)} /></Field>
                <Field label="Conduta seguinte"><TextArea rows={2} value={record.raciocinio.conduta} onChange={(v) => set("raciocinio.conduta", v)} /></Field>
              </div>
            </Section>
            <Section number={n + 7} title="Estudo de Caso (desta sessão)" open onToggle={() => {}}>
              <label className="flex items-center gap-2 cursor-pointer select-none" style={{ fontSize: 14, color: T.ink }}>
                <input type="checkbox" checked={!!record.estudoDeCasoSessao?.discutido} onChange={(e) => set("estudoDeCasoSessao.discutido", e.target.checked)} style={{ accentColor: T.accent, width: 16, height: 16 }} />
                Este atendimento foi discutido em estudo de caso com a supervisão
              </label>
              <Field label="Resumo da discussão (opcional)">
                <TextArea rows={2} value={record.estudoDeCasoSessao?.resumo || ""} onChange={(v) => set("estudoDeCasoSessao.resumo", v)} />
              </Field>
              <div style={{ fontSize: 11, color: T.inkFaint }}>Para um histórico de discussões mais longo e organizado por data, use o "Estudo de Caso" na ficha do paciente.</div>
            </Section>
            <Section number={n + 8} title="Validação da Supervisão" open onToggle={() => {}}>
              <div className="flex items-center gap-2 mb-1">
                <span style={{
                  fontSize: 11.5, fontWeight: 600, padding: "2px 8px", borderRadius: 999,
                  background: record.validacaoSupervisao?.status === "validado" ? T.accentSoft : T.warnSoft,
                  color: record.validacaoSupervisao?.status === "validado" ? T.accentDeep : T.warn,
                }}>
                  {record.validacaoSupervisao?.status === "validado" ? "✓ Validado" : "Pendente de validação"}
                </span>
                {record.validacaoSupervisao?.status === "validado" && (
                  <span style={{ fontSize: 11.5, color: T.inkFaint }}>
                    por {record.validacaoSupervisao.supervisorNome}{record.validacaoSupervisao.dataValidacao ? " em " + record.validacaoSupervisao.dataValidacao : ""}
                  </span>
                )}
              </div>
              <div className="grid sm:grid-cols-2 gap-3">
                <Field label="Supervisor(a) responsável"><TextInput value={record.validacaoSupervisao?.supervisorNome || ""} onChange={(v) => set("validacaoSupervisao.supervisorNome", v)} /></Field>
                <Field label="CRFa do(a) Supervisor(a)"><TextInput value={record.validacaoSupervisao?.crfaSupervisor || ""} onChange={(v) => set("validacaoSupervisao.crfaSupervisor", v)} /></Field>
              </div>
              <Field label="Comentário do(a) supervisor(a) (opcional)">
                <TextArea rows={2} value={record.validacaoSupervisao?.comentario || ""} onChange={(v) => set("validacaoSupervisao.comentario", v)} />
              </Field>
              {souSupervisor ? (
                record.validacaoSupervisao?.status === "validado" ? (
                  <button onClick={desfazerValidacao} className="self-start" style={{ fontSize: 12.5, color: T.warn }}>Desfazer validação</button>
                ) : (
                  <button onClick={validar} className="flex items-center gap-1.5 self-start" style={{ fontSize: 13, color: "white", background: T.accent, borderRadius: 7, padding: "7px 14px", fontWeight: 500 }}>
                    <Check size={14} /> Validar atendimento
                  </button>
                )
              ) : (
                <div style={{ fontSize: 11.5, color: T.inkFaint }}>Somente quem tem o papel "Supervisor(a)" ou "Administrador(a)" definido pelo administrador do sistema pode validar. Fale com um administrador se precisar desse acesso.</div>
              )}
            </Section>
          </>
        );
      })()}
    </div>
  );
}

/* ---------------------------------------------------------------------
   Case study (Estudo de Caso) — per-patient supervision notes
------------------------------------------------------------------------ */
function CaseStudyPanel({ patientName, notes, onAdd, onDelete }) {
  const [text, setText] = useState("");
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const sorted = [...notes].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  return (
    <div className="max-w-2xl mx-auto flex flex-col gap-5 pb-10">
      <div>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 21, fontWeight: 600, color: T.ink }}>Estudo de caso</div>
        <div style={{ fontSize: 13, color: T.inkSoft }}>{patientName} · anotações da discussão com a supervisão</div>
      </div>
      <div className="flex flex-col gap-2 p-4" style={{ background: T.surface, border: "1px solid " + T.line, borderRadius: 10 }}>
        <div className="grid sm:grid-cols-[160px_1fr] gap-3">
          <Field label="Data"><TextInput type="date" value={date} onChange={setDate} /></Field>
          <Field label="Anotação">
            <TextArea rows={4} value={text} onChange={setText} placeholder="O que foi discutido com a supervisão: hipóteses, condutas sugeridas, dúvidas, encaminhamentos…" />
          </Field>
        </div>
        <div className="flex justify-end">
          <button
            onClick={() => { if (text.trim()) { onAdd({ date, text }); setText(""); } }}
            className="flex items-center gap-1.5"
            style={{ fontSize: 13.5, color: "white", background: T.accent, borderRadius: 7, padding: "7px 14px", fontWeight: 500 }}
          >
            <Plus size={15} /> Adicionar anotação
          </button>
        </div>
      </div>
      <div className="flex flex-col gap-3">
        {sorted.length === 0 && <div style={{ color: T.inkFaint, fontSize: 13.5, textAlign: "center", padding: "24px 0" }}>Nenhuma anotação registrada ainda.</div>}
        {sorted.map((n) => (
          <div key={n.id} className="p-3 flex flex-col gap-1" style={{ background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 8 }}>
            <div className="flex items-center justify-between">
              <span style={{ fontSize: 12, color: T.inkFaint }}>{n.date}</span>
              <button onClick={() => onDelete(n.id)} style={{ color: T.warn }}><Trash2 size={14} /></button>
            </div>
            <div style={{ fontSize: 14, color: T.ink, whiteSpace: "pre-wrap" }}>{n.text}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   Progress report (Relatório de Acompanhamento) — timeline across areas
------------------------------------------------------------------------ */
function ProtocoloEvolucaoChart({ pontos }) {
  const validos = pontos.filter((p) => p.total != null);
  if (validos.length < 2) return null;
  const W = 420, H = 140, padL = 40, padR = 14, padT = 12, padB = 24;
  const plotW = W - padL - padR, plotH = H - padT - padB;
  const yMax = Math.max(...validos.map((p) => p.max || p.total), 1);
  const xFor = (i) => padL + (i / (validos.length - 1)) * plotW;
  const yFor = (v) => padT + plotH - (v / yMax) * plotH;
  const pathPts = validos.map((p, i) => [xFor(i), yFor(p.total)]);
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ maxWidth: 420, background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 6 }}>
      {[0, 0.5, 1].map((f) => (
        <g key={f}>
          <line x1={padL} x2={W - padR} y1={padT + plotH * (1 - f)} y2={padT + plotH * (1 - f)} stroke={T.lineSoft} strokeWidth="0.6" />
          <text x={padL - 4} y={padT + plotH * (1 - f) + 3} fontSize="8" textAnchor="end" fill={T.inkFaint}>{Math.round(yMax * f)}</text>
        </g>
      ))}
      <polyline points={pathPts.map((p) => p.join(",")).join(" ")} fill="none" stroke={T.accent} strokeWidth="1.6" />
      {pathPts.map(([x, y], i) => <circle key={i} cx={x} cy={y} r="3" fill={T.accent} />)}
      {validos.map((p, i) => <text key={i} x={xFor(i)} y={H - padB + 12} fontSize="7.5" textAnchor="middle" fill={T.inkFaint}>{p.data || "—"}</text>)}
    </svg>
  );
}
function ReportView({ patientName, sessions }) {
  const byArea = _.groupBy(sessions, "area");
  const sorted = [...sessions].sort((a, b) => (a.identificacao.data || "").localeCompare(b.identificacao.data || ""));
  const latest = sorted[sorted.length - 1];
  const protocolosPorNome = _.groupBy(
    sorted.filter((r) => r.clinico?.protocoloAplicado?.nome),
    (r) => r.clinico.protocoloAplicado.nome
  );
  return (
    <div className="max-w-3xl mx-auto flex flex-col gap-5 pb-10">
      <div>
        <div style={{ fontFamily: "'Fraunces', serif", fontSize: 21, fontWeight: 600, color: T.ink }}>Relatório de acompanhamento</div>
        <div style={{ fontSize: 13, color: T.inkSoft }}>{patientName} · {sessions.length} atendimento(s) registrado(s)</div>
      </div>
      <div className="grid sm:grid-cols-2 gap-3 p-4" style={{ background: T.surface, border: "1px solid " + T.line, borderRadius: 10 }}>
        <ReadRow label="Hipótese diagnóstica" value={latest?.paciente.hipotese} />
        <ReadRow label="CID" value={latest?.paciente.cid} />
        <ReadRow label="Nascimento" value={latest?.paciente.nascimento} />
        <ReadRow label="Responsável" value={latest?.paciente.responsavel} />
        <ReadRow label="Áreas acompanhadas" value={Object.keys(byArea).map((a) => AREAS[a]?.label).join(" · ")} />
        <ReadRow label="Período" value={sorted[0]?.identificacao.data ? `${sorted[0].identificacao.data} — ${latest.identificacao.data}` : ""} />
      </div>
      {Object.keys(protocolosPorNome).length > 0 && (
        <div className="flex flex-col gap-4">
          <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontWeight: 600, fontSize: 14, color: T.ink }}>Evolução dos Protocolos Aplicados</div>
          {Object.entries(protocolosPorNome).map(([nome, regs]) => {
            const pontos = regs.map((r) => {
              const t = computeProtocoloTotais(r.clinico.protocoloAplicado.itens);
              return { data: r.clinico.protocoloAplicado.dataAplicacao || r.identificacao.data, total: t?.total, max: t?.max };
            });
            return (
              <div key={nome} className="p-3 flex flex-col gap-2" style={{ background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 8 }}>
                <div style={{ fontSize: 13, fontWeight: 600, color: T.accentDeep }}>{nome}</div>
                <div style={{ fontSize: 11.5, color: T.inkFaint }}>{regs.length} aplicação(ões){pontos.length > 1 ? " — evolução do escore total ao longo do tempo:" : ""}</div>
                <ProtocoloEvolucaoChart pontos={pontos} />
                {pontos.length < 2 && pontos[0]?.total != null && (
                  <ReadRow label="Escore" value={`${pontos[0].total}${pontos[0].max ? ` / ${pontos[0].max}` : ""}`} />
                )}
              </div>
            );
          })}
        </div>
      )}
      <div className="flex flex-col gap-3">
        {sorted.map((r) => (
          <div key={r.id} className="p-4 flex flex-col gap-2" style={{ background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 8 }}>
            <div className="flex items-center justify-between flex-wrap gap-1">
              <span style={{ fontSize: 13.5, fontWeight: 600, color: T.accentDeep }}>{AREAS[r.area]?.label}</span>
              <span style={{ fontSize: 12, color: T.inkFaint }}>{r.identificacao.data || "sem data"}{r.identificacao.sessaoNum ? " · sessão " + r.identificacao.sessaoNum : ""}</span>
            </div>
            {r.procedimentos && <ReadRow label="Procedimentos" value={r.procedimentos} />}
            {r.evolucao && <ReadRow label="Evolução" value={r.evolucao} />}
            {r.plano && <ReadRow label="Plano seguinte" value={r.plano} />}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------
   App shell
------------------------------------------------------------------------ */
const blankPatient = (nome, slug, instituicao = "") => ({
  id: slug, nome, instituicao, nascimento: "", idade: "", responsavel: "", hipotese: "", cid: "", tempoAcompanhamento: "",
  anamnese: blankAnamneseGeral(), anamneseCompleta: false, criadoEm: null, atualizadoEm: null,
});
const blankPerfil = () => ({ nome: "", instituicao: "", crfa: "", papel: "estagiario" });

export default function App() {
  const [records, setRecords] = useState({});
  const [patients, setPatients] = useState({}); // slug -> paciente
  const [caseStudies, setCaseStudies] = useState({});
  const [loading, setLoading] = useState(true);
  const [selectedId, setSelectedId] = useState(null);
  const [selectedPatientSlug, setSelectedPatientSlug] = useState(null);
  const [mode, setMode] = useState("empty"); // empty|new-patient|patient|anamnese-edit|patient-edit|choose-area|new|edit|view|casestudy|report
  const [draft, setDraft] = useState(null);
  const [draftArea, setDraftArea] = useState(null);
  const [anamneseDraft, setAnamneseDraft] = useState(null);
  const [patientEditDraft, setPatientEditDraft] = useState(null);
  const [newPatientName, setNewPatientName] = useState("");
  const [query, setQuery] = useState("");
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState(false);
  const [mobileListOpen, setMobileListOpen] = useState(true);
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [pdfGenerating, setPdfGenerating] = useState(false);
  const [laudoGenerating, setLaudoGenerating] = useState(false);
  const [perfil, setPerfil] = useState(null); // null = ainda não carregado
  const [perfilDraft, setPerfilDraft] = useState(null);

  useEffect(() => {
    (async () => {
      try {
        let ids = [];
        try { const idx = await window.storage.get("record-index"); ids = idx ? JSON.parse(idx.value) : []; } catch (e) { ids = []; }
        const loaded = {};
        for (const id of ids) {
          try { const r = await window.storage.get("atendimento:" + id); if (r) loaded[id] = JSON.parse(r.value); } catch (e) {}
        }
        setRecords(loaded);

        let csIds = [];
        try { const idx = await window.storage.get("estudo-index"); csIds = idx ? JSON.parse(idx.value) : []; } catch (e) { csIds = []; }
        const csLoaded = {};
        for (const slug of csIds) {
          try { const r = await window.storage.get("estudo:" + slug); if (r) csLoaded[slug] = JSON.parse(r.value); } catch (e) {}
        }
        setCaseStudies(csLoaded);

        let pSlugs = [];
        try { const idx = await window.storage.get("patient-index"); pSlugs = idx ? JSON.parse(idx.value) : []; } catch (e) { pSlugs = []; }
        const pLoaded = {};
        for (const slug of pSlugs) {
          try { const r = await window.storage.get("paciente:" + slug); if (r) pLoaded[slug] = JSON.parse(r.value); } catch (e) {}
        }
        setPatients(pLoaded);

        try {
          const meu = await window.storage.get("meu-perfil", false);
          setPerfil(meu ? JSON.parse(meu.value) : blankPerfil());
        } catch (e) { setPerfil(blankPerfil()); }
      } finally { setLoading(false); }
    })();
  }, []);

  const savePerfil = async (novoPerfil) => {
    try { await window.storage.set("meu-perfil", JSON.stringify(novoPerfil), false); } catch (e) {}
    setPerfil(novoPerfil);
    setPerfilDraft(null);
  };

  const patientGroups = useMemo(() => {
    const groups = {};
    Object.values(records).forEach((r) => {
      const name = (r.paciente.nome || "Sem nome").trim();
      if (!groups[name]) groups[name] = [];
      groups[name].push(r);
    });
    Object.values(groups).forEach((list) => list.sort((a, b) => (b.identificacao.data || "").localeCompare(a.identificacao.data || "")));
    return groups;
  }, [records]);

  // Unified patient rows: formal Paciente entities + legacy names inferred from atendimentos.
  const patientRows = useMemo(() => {
    const slugs = new Set(Object.keys(patients));
    Object.keys(patientGroups).forEach((n) => { if (n !== "Sem nome") slugs.add(slugify(n)); });
    const minhaInstituicao = (perfil?.instituicao || "").trim().toLowerCase();
    const rows = Array.from(slugs)
      .map((slug) => {
        const p = patients[slug];
        const legacyName = Object.keys(patientGroups).find((n) => slugify(n) === slug);
        const nome = p?.nome || legacyName || slug;
        const sessions = patientGroups[nome] || [];
        return { slug, nome, instituicao: p?.instituicao || "", sessionsCount: sessions.length, lastDate: sessions[0]?.identificacao.data || "" };
      })
      .filter((r) => !minhaInstituicao || !r.instituicao || r.instituicao.trim().toLowerCase() === minhaInstituicao);
    rows.sort((a, b) => a.nome.localeCompare(b.nome, "pt-BR"));
    return rows;
  }, [patients, patientGroups, perfil]);

  const filteredPatientRows = useMemo(() => {
    if (!query.trim()) return patientRows;
    const q = query.trim().toLowerCase();
    return patientRows.filter((r) => r.nome.toLowerCase().includes(q));
  }, [patientRows, query]);

  const persistIndex = async (ids) => { try { await window.storage.set("record-index", JSON.stringify(ids)); } catch (e) {} };
  const persistCaseIndex = async (slugs) => { try { await window.storage.set("estudo-index", JSON.stringify(slugs)); } catch (e) {} };
  const persistPatientIndex = async (slugs) => { try { await window.storage.set("patient-index", JSON.stringify(slugs)); } catch (e) {} };

  const currentPatient = selectedPatientSlug ? patients[selectedPatientSlug] : null;
  const currentPatientSessions = currentPatient ? (patientGroups[currentPatient.nome] || []) : [];

  const savePatientData = async (slug, updater) => {
    const base = patients[slug] || blankPatient(slug, slug);
    const updated = typeof updater === "function" ? updater(base) : updater;
    const now = new Date().toISOString();
    const withMeta = { ...updated, criadoEm: updated.criadoEm || now, atualizadoEm: now };
    try {
      await window.storage.set("paciente:" + slug, JSON.stringify(withMeta));
      const slugs = Array.from(new Set([...Object.keys(patients), slug]));
      await persistPatientIndex(slugs);
      setPatients((prev) => ({ ...prev, [slug]: withMeta }));
    } catch (e) {}
  };

  const createPatient = async () => {
    const nome = newPatientName.trim();
    if (!nome) return;
    const slug = slugify(nome);
    if (!patients[slug]) await savePatientData(slug, blankPatient(nome, slug, perfil?.instituicao || ""));
    setNewPatientName("");
    setSelectedPatientSlug(slug);
    setMode("patient");
    setMobileListOpen(false);
  };

  const openPatient = (slug) => {
    if (!patients[slug]) {
      const legacyName = Object.keys(patientGroups).find((n) => slugify(n) === slug);
      if (legacyName) {
        const latest = (patientGroups[legacyName] || [])[0];
        const seed = blankPatient(legacyName, slug);
        if (latest) {
          seed.nascimento = latest.paciente.nascimento || "";
          seed.idade = latest.paciente.idade || "";
          seed.responsavel = latest.paciente.responsavel || "";
          seed.hipotese = latest.paciente.hipotese || "";
          seed.cid = latest.paciente.cid || "";
          seed.tempoAcompanhamento = latest.paciente.tempoAcompanhamento || "";
        }
        setPatients((prev) => ({ ...prev, [slug]: seed }));
      }
    }
    setSelectedPatientSlug(slug);
    setSelectedId(null);
    setMode("patient");
    setMobileListOpen(false);
  };

  const startAnamnese = () => { setAnamneseDraft(_.cloneDeep(currentPatient.anamnese)); setMode("anamnese-edit"); };
  const saveAnamnese = async () => {
    await savePatientData(selectedPatientSlug, (prev) => ({ ...prev, anamnese: anamneseDraft, anamneseCompleta: true }));
    setAnamneseDraft(null);
    setMode("patient");
  };

  const startEditPatientInfo = () => setPatientEditDraft({ ..._.pick(currentPatient, ["nome", "instituicao", "nascimento", "idade", "responsavel", "hipotese", "cid", "tempoAcompanhamento"]) });
  const savePatientInfo = async () => {
    await savePatientData(selectedPatientSlug, (prev) => ({ ...prev, ...patientEditDraft }));
    setPatientEditDraft(null);
  };

  const startChooseArea = () => { setMode("choose-area"); setMobileListOpen(false); };
  const startNewWithArea = (areaId) => {
    if (!currentPatient) return;
    const blank = blankRecord(areaId);
    blank.pacienteSlug = selectedPatientSlug;
    blank.paciente = {
      nome: currentPatient.nome, nascimento: currentPatient.nascimento, idade: currentPatient.idade,
      responsavel: currentPatient.responsavel, hipotese: currentPatient.hipotese, cid: currentPatient.cid,
      tempoAcompanhamento: currentPatient.tempoAcompanhamento,
    };
    if (perfil?.instituicao) blank.identificacao.instituicao = perfil.instituicao;
    if (perfil?.nome) blank.identificacao.supervisor = perfil.nome;
    if (perfil?.crfa) blank.identificacao.crfa = perfil.crfa;
    const sameAreaCount = currentPatientSessions.filter((r) => r.area === areaId).length;
    blank.identificacao.sessaoNum = String(sameAreaCount + 1);
    setDraftArea(areaId);
    setDraft(blank);
    setMode("new");
    setSelectedId(null);
  };

  const openView = (id) => {
    setSelectedId(id);
    const rec = records[id];
    if (rec) setSelectedPatientSlug(slugify(rec.paciente.nome));
    setMode("view");
    setMobileListOpen(false);
  };
  const startEdit = (id) => {
    const base = blankRecord(records[id].area);
    setDraft(_.merge({}, base, _.cloneDeep(records[id])));
    setDraftArea(records[id].area);
    setSelectedId(id);
    setMode("edit");
  };
  const cancelEdit = () => {
    setDraft(null); setDraftArea(null);
    if (selectedId) setMode("view");
    else if (selectedPatientSlug) setMode("patient");
    else setMode("empty");
  };

  const save = async () => {
    if (!draft.paciente.nome.trim()) { setSaveError(true); return; }
    setSaveError(false); setSaving(true);
    const now = new Date().toISOString();
    const id = draft.id || genId();
    const toSave = { ...draft, id, criadoEm: draft.criadoEm || now, atualizadoEm: now };
    try {
      await window.storage.set("atendimento:" + id, JSON.stringify(toSave));
      const newIds = Array.from(new Set([...Object.keys(records), id]));
      await persistIndex(newIds);
      setRecords((prev) => ({ ...prev, [id]: toSave }));
      setSelectedId(id); setMode("view"); setDraft(null); setDraftArea(null);
    } catch (e) { setSaveError(true); } finally { setSaving(false); }
  };

  const doDelete = async (id) => {
    try {
      await window.storage.delete("atendimento:" + id);
      const remaining = Object.keys(records).filter((k) => k !== id);
      await persistIndex(remaining);
      const next = { ...records }; delete next[id]; setRecords(next);
      if (selectedId === id) { setSelectedId(null); setMode(selectedPatientSlug ? "patient" : "empty"); }
    } catch (e) {}
    setConfirmDelete(null);
  };

  const openCaseStudy = () => { setMode("casestudy"); setMobileListOpen(false); };
  const openReport = () => { setMode("report"); setMobileListOpen(false); };

  const addCaseNote = async (name, note) => {
    const slug = slugify(name);
    const existing = caseStudies[slug] || { name, notes: [] };
    const updated = { name, notes: [...existing.notes, { id: genId(), createdAt: new Date().toISOString(), ...note }] };
    try {
      await window.storage.set("estudo:" + slug, JSON.stringify(updated));
      const slugs = Array.from(new Set([...Object.keys(caseStudies), slug]));
      await persistCaseIndex(slugs);
      setCaseStudies((prev) => ({ ...prev, [slug]: updated }));
    } catch (e) {}
  };
  const deleteCaseNote = async (name, noteId) => {
    const slug = slugify(name);
    const existing = caseStudies[slug];
    if (!existing) return;
    const updated = { ...existing, notes: existing.notes.filter((n) => n.id !== noteId) };
    try { await window.storage.set("estudo:" + slug, JSON.stringify(updated)); setCaseStudies((prev) => ({ ...prev, [slug]: updated })); } catch (e) {}
  };

  const doPrint = () => window.print();
  const handleDownloadPDF = async (record) => {
    setPdfGenerating(true);
    try { await generateRecordPDF(record); } catch (e) { window.print(); } finally { setPdfGenerating(false); }
  };
  const handleDownloadLaudo = async (record) => {
    setLaudoGenerating(true);
    try { await generateLaudoPDF(record); } catch (e) { window.alert("Não foi possível gerar o laudo em PDF agora. Tente novamente em instantes."); } finally { setLaudoGenerating(false); }
  };
  const current = selectedId ? records[selectedId] : null;

  const exportSpreadsheet = (groupBy) => {
    const rows = Object.values(records).map((r) => ({
      "Fonoaudiólogo(a) responsável": r.identificacao.supervisor,
      "Estagiário(a)": r.identificacao.estagiario,
      Paciente: r.paciente.nome,
      "Área": AREAS[r.area]?.label || r.area,
      Data: r.identificacao.data,
      "Sessão nº": r.identificacao.sessaoNum,
      "Protocolo aplicado": r.clinico?.protocoloAplicado?.nome || "",
      "Hipótese diagnóstica": r.paciente.hipotese,
      "CID": r.paciente.cid,
      "Procedimentos realizados": r.procedimentos,
      "Evolução clínica": r.evolucao,
      "Orientações à família": r.orientacoes,
      "Plano para próxima sessão": r.plano,
    }));
    const sortKey = groupBy === "estagiario" ? "Fonoaudiólogo(a) responsável" : "Paciente";
    rows.sort((a, b) => (a[sortKey] || "").localeCompare(b[sortKey] || "", "pt-BR") || (a.Data || "").localeCompare(b.Data || ""));
    if (rows.length === 0) return;
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = Object.keys(rows[0]).map((k) => ({ wch: Math.min(40, Math.max(12, k.length + 4)) }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, groupBy === "estagiario" ? "Por profissional" : "Por paciente");
    XLSX.writeFile(wb, groupBy === "estagiario" ? "atendimentos_por_profissional.xlsx" : "atendimentos_por_paciente.xlsx");
  };

  if (perfil === null) {
    return <div style={{ padding: 40, fontFamily: "'IBM Plex Sans', sans-serif", color: T.inkSoft }}>Carregando…</div>;
  }

  if (!perfil.instituicao || perfilDraft) {
    const draft = perfilDraft || blankPerfil();
    return (
      <div style={{ background: T.paper, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'IBM Plex Sans', sans-serif", padding: 16 }}>
        <div style={{ width: "100%", maxWidth: 400, background: T.surface, border: "1px solid " + T.line, borderRadius: 12, padding: 28, display: "flex", flexDirection: "column", gap: 14 }}>
          <div>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 20, fontWeight: 600, color: T.ink }}>
              {perfil.instituicao ? "Editar perfil profissional" : "Configure seu perfil"}
            </div>
            <div style={{ fontSize: 12.5, color: T.inkSoft }}>
              A instituição escolhida define quais pacientes aparecem para você — pacientes da mesma instituição ficam visíveis para todos os profissionais que os atendem.
            </div>
          </div>
          <Field label="Seu nome"><TextInput value={draft.nome} onChange={(v) => setPerfilDraft({ ...draft, nome: v })} placeholder="Nome completo" /></Field>
          {typeof window !== "undefined" && window.authProfile ? (
            <Field label="Seu papel na equipe">
              <div style={{ ...inputBase, background: T.surfaceSoft, color: T.inkSoft }}>
                {{ estagiario: "Estagiário(a)", fonoaudiologo: "Fonoaudiólogo(a)", supervisor: "Supervisor(a)", admin: "Administrador(a)" }[window.authProfile.papel] || "Estagiário(a)"}
              </div>
              <span style={{ fontSize: 10.5, color: T.inkFaint }}>Definido por um administrador do sistema — fale com ele(a) para alterar.</span>
            </Field>
          ) : (
            <Field label="Seu papel na equipe">
              <select value={draft.papel || "estagiario"} onChange={(e) => setPerfilDraft({ ...draft, papel: e.target.value })} style={inputBase}>
                <option value="estagiario">Estagiário(a)</option>
                <option value="fonoaudiologo">Fonoaudiólogo(a)</option>
                <option value="supervisor">Supervisor(a)</option>
              </select>
            </Field>
          )}
          <Field label="Instituição"><TextInput value={draft.instituicao} onChange={(v) => setPerfilDraft({ ...draft, instituicao: v })} placeholder="Ex.: Centro Universitário Nobre – UNIFAN" /></Field>
          <Field label="CRFa (se houver)"><TextInput value={draft.crfa} onChange={(v) => setPerfilDraft({ ...draft, crfa: v })} /></Field>
          <div className="flex justify-end gap-3">
            {perfil.instituicao && <button onClick={() => setPerfilDraft(null)} style={{ fontSize: 13.5, color: T.inkSoft }}>Cancelar</button>}
            <button
              onClick={() => draft.instituicao.trim() && savePerfil(draft)}
              style={{ background: T.accent, color: "white", border: "none", borderRadius: 8, padding: "9px 16px", fontSize: 14, fontWeight: 500, cursor: "pointer" }}
            >
              Salvar e continuar
            </button>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ background: T.paper, minHeight: "100%", fontFamily: "'IBM Plex Sans', sans-serif" }}>
      <style>{`
        ${FONT_IMPORT}
        * { box-sizing: border-box; }
        ::selection { background: ${T.accentSoft}; }
        @media print {
          .no-print { display: none !important; }
          #app-shell { background: white !important; }
          #print-only { display: block !important; }
        }
        #print-only { display: none; }
      `}</style>

      <div id="app-shell" className="flex flex-col lg:flex-row" style={{ minHeight: "100%" }}>
        {/* Sidebar — lista de pacientes */}
        <div className={"no-print flex-col " + (mobileListOpen ? "flex" : "hidden lg:flex")}
          style={{ width: "100%", maxWidth: 320, minWidth: 280, borderRight: "1px solid " + T.line, background: T.surface }}>
          <div className="px-4 pt-5 pb-3" style={{ borderBottom: "1px solid " + T.lineSoft }}>
            <div className="flex items-center gap-2">
              <div style={{ width: 30, height: 30, borderRadius: 7, background: T.accent, display: "flex", alignItems: "center", justifyContent: "center" }}>
                <ClipboardList size={16} color="white" />
              </div>
              <div>
                <div style={{ fontFamily: "'Fraunces', serif", fontSize: 15.5, fontWeight: 600, color: T.ink, lineHeight: 1.1 }}>Protocolo de Atendimento</div>
                <div style={{ fontSize: 11.5, color: T.inkSoft }}>Fonoaudiologia · ficha do paciente</div>
              </div>
            </div>
            <div className="flex items-center justify-between mt-3" style={{ fontSize: 11.5, color: T.inkSoft }}>
              <span>Instituição: <strong style={{ color: T.ink }}>{perfil.instituicao}</strong></span>
              <button onClick={() => setPerfilDraft({ ...perfil })} style={{ color: T.accent, fontWeight: 500 }}>trocar</button>
            </div>
            <button onClick={() => setMode("new-patient")} className="w-full mt-3 flex items-center justify-center gap-2"
              style={{ background: T.accent, color: "white", borderRadius: 8, padding: "9px 0", fontSize: 14, fontWeight: 500 }}>
              <Plus size={16} /> Novo paciente
            </button>
          </div>

          <div className="px-4 py-3 flex flex-col gap-2">
            <div className="flex items-center gap-2 px-3 py-2" style={{ background: T.surfaceSoft, borderRadius: 8, border: "1px solid " + T.lineSoft }}>
              <Search size={15} color={T.inkFaint} />
              <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar paciente"
                style={{ border: "none", outline: "none", background: "transparent", fontSize: 13.5, flex: 1, color: T.ink }} />
            </div>
            {Object.keys(records).length > 0 && (
              <div className="flex gap-2">
                <button onClick={() => exportSpreadsheet("estagiario")}
                  className="flex-1 flex items-center justify-center gap-1.5 px-2 py-1.5"
                  style={{ background: T.surfaceSoft, border: "1px solid " + T.lineSoft, borderRadius: 7, fontSize: 11.5, color: T.accentDeep, fontWeight: 500 }}
                  title="Baixar planilha .xlsx com todos os atendimentos agrupados por profissional">
                  <FileSpreadsheet size={13} /> Por profissional
                </button>
                <button onClick={() => exportSpreadsheet("paciente")}
                  className="flex-1 flex items-center justify-center gap-1.5 px-2 py-1.5"
                  style={{ background: T.surfaceSoft, border: "1px solid " + T.lineSoft, borderRadius: 7, fontSize: 11.5, color: T.accentDeep, fontWeight: 500 }}
                  title="Baixar planilha .xlsx com todos os atendimentos agrupados por paciente">
                  <FileSpreadsheet size={13} /> Por paciente
                </button>
              </div>
            )}
          </div>

          <div className="flex-1 overflow-y-auto px-2 pb-6">
            {loading ? (
              <div className="px-3 py-6 text-center" style={{ color: T.inkFaint, fontSize: 13 }}>Carregando…</div>
            ) : filteredPatientRows.length === 0 ? (
              <div className="px-3 py-6 text-center" style={{ color: T.inkFaint, fontSize: 13 }}>
                {patientRows.length === 0 ? "Nenhum paciente cadastrado ainda." : "Nenhum paciente encontrado."}
              </div>
            ) : (
              filteredPatientRows.map((row) => (
                <button key={row.slug} onClick={() => openPatient(row.slug)}
                  className="w-full flex items-center gap-2 px-3 py-2.5 text-left mb-0.5"
                  style={{ borderRadius: 7, background: selectedPatientSlug === row.slug ? T.accentSoft : "transparent" }}>
                  <User size={14} color={T.inkFaint} />
                  <div className="flex-1" style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 13.5, color: T.ink, fontWeight: 500, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{row.nome}</div>
                    <div style={{ fontSize: 11, color: T.inkFaint }}>{row.sessionsCount} atendimento(s){row.lastDate ? " · último em " + row.lastDate : ""}</div>
                  </div>
                </button>
              ))
            )}
          </div>
        </div>

        {/* Main panel */}
        <div className="flex-1 flex flex-col" style={{ minWidth: 0 }}>
          <div className="no-print flex items-center gap-3 px-4 lg:px-8 py-3" style={{ borderBottom: "1px solid " + T.line, background: T.surface }}>
            <button className="lg:hidden flex items-center gap-1" onClick={() => setMobileListOpen((s) => !s)} style={{ fontSize: 13, color: T.accent }}>
              <ArrowLeft size={15} /> Lista
            </button>
            {mode !== "patient" && selectedPatientSlug && currentPatient && (
              <button onClick={() => setMode("patient")} className="flex items-center gap-1" style={{ fontSize: 13, color: T.accent }}>
                <ArrowLeft size={15} /> Ficha de {currentPatient.nome}
              </button>
            )}
            <div style={{ flex: 1 }} />
            {mode === "view" && current && (
              <>
                <button onClick={() => startEdit(current.id)} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: T.ink }}><Pencil size={15} /> Editar</button>
                <button onClick={() => handleDownloadPDF(current)} disabled={pdfGenerating} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: T.ink }} title="Baixa um PDF do atendimento pronto para colher as assinaturas do estagiário e do supervisor"><Printer size={15} /> {pdfGenerating ? "Gerando PDF…" : "Baixar PDF"}</button>
                {isAudiologyRecord(current) && (
                  <button onClick={() => handleDownloadLaudo(current)} disabled={laudoGenerating} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: T.accentDeep }} title="Baixa um laudo audiológico curto (achados, conclusão e assinaturas), separado do registro completo do atendimento">
                    <FileBarChart size={15} /> {laudoGenerating ? "Gerando laudo…" : "Baixar Laudo (PDF)"}
                  </button>
                )}
                <button onClick={() => setConfirmDelete(current.id)} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: T.warn }}><Trash2 size={15} /> Excluir</button>
              </>
            )}
            {(mode === "new" || mode === "edit") && (
              <>
                <button onClick={cancelEdit} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: T.inkSoft }}><X size={15} /> Cancelar</button>
                <button onClick={save} disabled={saving} className="flex items-center gap-1.5"
                  style={{ fontSize: 13.5, color: "white", background: T.accent, borderRadius: 7, padding: "7px 14px", fontWeight: 500 }}>
                  <Save size={15} /> {saving ? "Salvando…" : "Salvar"}
                </button>
              </>
            )}
            {mode === "anamnese-edit" && (
              <>
                <button onClick={() => { setAnamneseDraft(null); setMode("patient"); }} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: T.inkSoft }}><X size={15} /> Cancelar</button>
                <button onClick={saveAnamnese} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: "white", background: T.accent, borderRadius: 7, padding: "7px 14px", fontWeight: 500 }}>
                  <Save size={15} /> Salvar anamnese
                </button>
              </>
            )}
            {mode === "report" && (
              <button onClick={doPrint} className="flex items-center gap-1.5" style={{ fontSize: 13.5, color: T.ink }}><Printer size={15} /> Imprimir</button>
            )}
          </div>

          <div className="flex-1 overflow-y-auto px-4 lg:px-8 py-6">
            {mode === "empty" && (
              <div className="h-full flex flex-col items-center justify-center text-center gap-3 py-24">
                <ClipboardList size={34} color={T.inkFaint} />
                <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19, color: T.ink }}>Selecione um paciente ou cadastre um novo</div>
                <div style={{ fontSize: 13.5, color: T.inkSoft, maxWidth: 380 }}>
                  Cada paciente tem uma ficha própria: anamnese única, histórico de todos os atendimentos
                  (de qualquer área) e dos protocolos aplicados.
                </div>
                <button onClick={() => setMode("new-patient")} className="mt-2 flex items-center gap-2"
                  style={{ background: T.accent, color: "white", borderRadius: 8, padding: "9px 16px", fontSize: 14, fontWeight: 500 }}>
                  <Plus size={16} /> Novo paciente
                </button>
              </div>
            )}

            {mode === "new-patient" && (
              <div className="max-w-md mx-auto flex flex-col gap-4 py-10">
                <div>
                  <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19, color: T.ink }}>Novo paciente</div>
                  <div style={{ fontSize: 13.5, color: T.inkSoft }}>Depois de criar, você preenche a anamnese e escolhe a área do primeiro atendimento.</div>
                </div>
                <Field label="Nome do paciente">
                  <TextInput value={newPatientName} onChange={setNewPatientName} placeholder="Nome completo" />
                </Field>
                <div style={{ fontSize: 12, color: T.inkFaint }}>Será cadastrado em: <strong style={{ color: T.inkSoft }}>{perfil.instituicao}</strong></div>
                <div className="flex justify-end gap-3">
                  <button onClick={() => { setMode("empty"); setNewPatientName(""); }} style={{ fontSize: 13.5, color: T.inkSoft }}>Cancelar</button>
                  <button onClick={createPatient} className="flex items-center gap-1.5" style={{ fontSize: 14, color: "white", background: T.accent, borderRadius: 8, padding: "9px 18px", fontWeight: 500 }}>
                    <Plus size={15} /> Criar paciente
                  </button>
                </div>
              </div>
            )}

            {mode === "patient" && currentPatient && (
              <div className="max-w-3xl mx-auto flex flex-col gap-5 pb-10">
                <div className="flex items-start justify-between gap-3">
                  <div>
                    <div style={{ fontFamily: "'Fraunces', serif", fontSize: 22, fontWeight: 600, color: T.ink }}>{currentPatient.nome}</div>
                    <div style={{ fontSize: 13, color: T.inkSoft }}>
                      {currentPatient.nascimento ? "Nascimento: " + currentPatient.nascimento : "Nascimento não informado"}
                      {currentPatient.idade ? " · " + currentPatient.idade : ""}
                    </div>
                  </div>
                  <button onClick={startEditPatientInfo} className="flex items-center gap-1" style={{ fontSize: 12.5, color: T.accent }}>
                    <Pencil size={13} /> Editar dados
                  </button>
                </div>

                {patientEditDraft && (
                  <div className="p-4 flex flex-col gap-3" style={{ background: T.surface, border: "1px solid " + T.line, borderRadius: 10 }}>
                    <div className="grid sm:grid-cols-2 gap-3">
                      <Field label="Nome" span={2}><TextInput value={patientEditDraft.nome} onChange={(v) => setPatientEditDraft((p) => ({ ...p, nome: v }))} /></Field>
                      <Field label="Instituição" span={2}><TextInput value={patientEditDraft.instituicao} onChange={(v) => setPatientEditDraft((p) => ({ ...p, instituicao: v }))} /></Field>
                      <Field label="Nascimento"><TextInput type="date" value={patientEditDraft.nascimento} onChange={(v) => setPatientEditDraft((p) => ({ ...p, nascimento: v }))} /></Field>
                      <Field label="Idade"><TextInput value={patientEditDraft.idade} onChange={(v) => setPatientEditDraft((p) => ({ ...p, idade: v }))} /></Field>
                      <Field label="Responsável" span={2}><TextInput value={patientEditDraft.responsavel} onChange={(v) => setPatientEditDraft((p) => ({ ...p, responsavel: v }))} /></Field>
                      <Field label="Hipótese diagnóstica" span={2}><TextInput value={patientEditDraft.hipotese} onChange={(v) => setPatientEditDraft((p) => ({ ...p, hipotese: v }))} /></Field>
                      <Field label="CID"><TextInput value={patientEditDraft.cid} onChange={(v) => setPatientEditDraft((p) => ({ ...p, cid: v }))} /></Field>
                      <Field label="Tempo de acompanhamento"><TextInput value={patientEditDraft.tempoAcompanhamento} onChange={(v) => setPatientEditDraft((p) => ({ ...p, tempoAcompanhamento: v }))} /></Field>
                    </div>
                    <div className="flex justify-end gap-3">
                      <button onClick={() => setPatientEditDraft(null)} style={{ fontSize: 13, color: T.inkSoft }}>Cancelar</button>
                      <button onClick={savePatientInfo} style={{ fontSize: 13, color: "white", background: T.accent, borderRadius: 7, padding: "6px 12px" }}>Salvar</button>
                    </div>
                  </div>
                )}

                <div className="flex flex-wrap gap-3 p-3" style={{ background: T.surfaceSoft, borderRadius: 10 }}>
                  <button onClick={startAnamnese} className="flex items-center gap-1.5" style={{ fontSize: 13, color: T.accentDeep, fontWeight: 500 }}>
                    <ClipboardList size={15} /> {currentPatient.anamneseCompleta ? "Ver / editar anamnese" : "Preencher anamnese"}
                  </button>
                  <button onClick={openCaseStudy} className="flex items-center gap-1.5" style={{ fontSize: 13, color: T.accent }}>
                    <BookOpen size={15} /> Estudo de caso
                  </button>
                  <button onClick={openReport} className="flex items-center gap-1.5" style={{ fontSize: 13, color: T.accent }}>
                    <FileBarChart size={15} /> Relatório de acompanhamento
                  </button>
                  {!currentPatient.anamneseCompleta && (
                    <span className="flex items-center gap-1" style={{ fontSize: 12, color: T.warn }}>
                      <AlertCircle size={13} /> Anamnese ainda não preenchida
                    </span>
                  )}
                </div>

                <button onClick={startChooseArea} className="flex items-center justify-center gap-2 py-3"
                  style={{ background: T.accent, color: "white", borderRadius: 10, fontSize: 14.5, fontWeight: 500 }}>
                  <Plus size={17} /> Novo atendimento
                </button>

                <div>
                  <div style={{ fontFamily: "'IBM Plex Sans', sans-serif", fontWeight: 600, fontSize: 14, color: T.ink, marginBottom: 8 }}>
                    Histórico de atendimentos ({currentPatientSessions.length})
                  </div>
                  {currentPatientSessions.length === 0 ? (
                    <div style={{ fontSize: 13, color: T.inkFaint, textAlign: "center", padding: "20px 0" }}>Nenhum atendimento registrado ainda para este paciente.</div>
                  ) : (
                    <div className="flex flex-col gap-2">
                      {currentPatientSessions.map((r) => (
                        <button key={r.id} onClick={() => openView(r.id)} className="flex items-center gap-3 p-3 text-left"
                          style={{ background: T.surface, border: "1px solid " + T.lineSoft, borderRadius: 8 }}>
                          <Calendar size={15} color={T.inkFaint} />
                          <div className="flex-1" style={{ minWidth: 0 }}>
                            <div className="flex items-center gap-2 flex-wrap">
                              <span style={{ fontSize: 13.5, color: T.ink, fontWeight: 500 }}>
                                {SHORT_LABELS[r.area] || AREAS[r.area]?.label}
                                {r.identificacao.sessaoNum ? " · sessão " + r.identificacao.sessaoNum : ""}
                              </span>
                              <span style={{
                                fontSize: 10, fontWeight: 600, padding: "1px 7px", borderRadius: 999,
                                background: r.validacaoSupervisao?.status === "validado" ? T.accentSoft : T.warnSoft,
                                color: r.validacaoSupervisao?.status === "validado" ? T.accentDeep : T.warn,
                              }}>
                                {r.validacaoSupervisao?.status === "validado" ? "✓ Validado" : "Pendente"}
                              </span>
                              {r.estudoDeCasoSessao?.discutido && (
                                <span className="flex items-center gap-0.5" style={{ fontSize: 10, color: T.accent }} title="Discutido em estudo de caso">
                                  <BookOpen size={11} />
                                </span>
                              )}
                            </div>
                            <div style={{ fontSize: 11.5, color: T.inkFaint }}>
                              {r.identificacao.data || "sem data"}
                              {r.identificacao.supervisor ? " · Fono: " + r.identificacao.supervisor : ""}
                              {r.identificacao.estagiario ? " · Estagiário(a): " + r.identificacao.estagiario : ""}
                              {r.clinico?.protocoloAplicado?.nome ? " · Protocolo: " + r.clinico.protocoloAplicado.nome : ""}
                            </div>
                          </div>
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}

            {mode === "anamnese-edit" && anamneseDraft && currentPatient && (
              <div className="max-w-3xl mx-auto flex flex-col gap-4">
                <div>
                  <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19, color: T.ink }}>Anamnese — {currentPatient.nome}</div>
                  <div style={{ fontSize: 13, color: T.inkSoft }}>Preenchida uma única vez; edite aqui sempre que precisar atualizar.</div>
                </div>
                <ClinicoForm sections={ANAMNESE_GERAL_SECTIONS} startNumber={1} clinico={anamneseDraft} setClinico={setAnamneseDraft} />
                <div className="flex justify-end gap-3 pb-8">
                  <button onClick={() => { setAnamneseDraft(null); setMode("patient"); }} style={{ fontSize: 13.5, color: T.inkSoft, padding: "8px 14px" }}>Cancelar</button>
                  <button onClick={saveAnamnese} className="flex items-center gap-1.5" style={{ fontSize: 14, color: "white", background: T.accent, borderRadius: 8, padding: "9px 18px", fontWeight: 500 }}>
                    <Save size={15} /> Salvar anamnese
                  </button>
                </div>
              </div>
            )}

            {mode === "choose-area" && currentPatient && (
              <div className="max-w-2xl mx-auto flex flex-col gap-4 py-6">
                <div>
                  <div style={{ fontFamily: "'Fraunces', serif", fontSize: 19, color: T.ink }}>Atendimento de {currentPatient.nome} — qual área?</div>
                  <div style={{ fontSize: 13.5, color: T.inkSoft }}>O protocolo se ajusta aos campos clínicos dessa área. Os dados do paciente não precisam ser repetidos.</div>
                </div>
                <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3">
                  {AREA_ORDER.map((id) => {
                    const Icon = AREAS[id].icon;
                    return (
                      <button key={id} onClick={() => startNewWithArea(id)}
                        className="flex items-start gap-3 p-4 text-left"
                        style={{ background: T.surface, border: "1px solid " + T.line, borderRadius: 10 }}>
                        <div style={{ width: 34, height: 34, borderRadius: 8, background: T.accentSoft, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                          <Icon size={17} color={T.accentDeep} />
                        </div>
                        <div>
                          <div style={{ fontSize: 14, fontWeight: 600, color: T.ink }}>{AREAS[id].label}</div>
                          <div style={{ fontSize: 12, color: T.inkFaint }}>{AREAS[id].disciplina}</div>
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {(mode === "new" || mode === "edit") && draft && (
              <div className="max-w-3xl mx-auto flex flex-col gap-4">
                {saveError && (
                  <div className="flex items-center gap-2 px-3 py-2" style={{ background: T.warnSoft, color: T.warn, borderRadius: 7, fontSize: 13 }}>
                    <AlertCircle size={15} /> Informe ao menos o nome do paciente para salvar.
                  </div>
                )}
                <RecordForm areaId={draftArea} record={draft} setRecord={setDraft} perfil={perfil} />
                <div className="flex justify-end gap-3 pb-8">
                  <button onClick={cancelEdit} style={{ fontSize: 13.5, color: T.inkSoft, padding: "8px 14px" }}>Cancelar</button>
                  <button onClick={save} disabled={saving} className="flex items-center gap-1.5"
                    style={{ fontSize: 14, color: "white", background: T.accent, borderRadius: 8, padding: "9px 18px", fontWeight: 500 }}>
                    <Save size={15} /> {saving ? "Salvando…" : "Salvar atendimento"}
                  </button>
                </div>
              </div>
            )}

            {mode === "view" && current && (
              <div className="max-w-3xl mx-auto pb-8">
                <div className="mb-5">
                  <div style={{ fontFamily: "'Fraunces', serif", fontSize: 21, fontWeight: 600, color: T.ink }}>{current.paciente.nome}</div>
                  <div style={{ fontSize: 13, color: T.inkSoft }}>
                    {AREAS[current.area]?.label} · {current.identificacao.data || "sem data"}
                    {current.identificacao.sessaoNum ? " · Sessão nº " + current.identificacao.sessaoNum : ""}
                  </div>
                </div>
                <RecordReadView record={current} />
              </div>
            )}

            {mode === "casestudy" && currentPatient && (
              <CaseStudyPanel
                patientName={currentPatient.nome}
                notes={(caseStudies[slugify(currentPatient.nome)] || { notes: [] }).notes}
                onAdd={(note) => addCaseNote(currentPatient.nome, note)}
                onDelete={(id) => deleteCaseNote(currentPatient.nome, id)}
              />
            )}

            {mode === "report" && currentPatient && (
              <ReportView patientName={currentPatient.nome} sessions={currentPatientSessions} />
            )}
          </div>
        </div>
      </div>

      {/* Print-only clean copy */}
      <div id="print-only">
        {mode === "view" && current && (
          <div style={{ padding: 24 }}>
            <div style={{ textAlign: "center", marginBottom: 4, fontFamily: "'Fraunces', serif", fontWeight: 600, fontSize: 16 }}>{current.identificacao.instituicao}</div>
            <div style={{ textAlign: "center", fontSize: 12, marginBottom: 14 }}>{current.identificacao.curso} — {current.identificacao.disciplina}</div>
            <div style={{ textAlign: "center", fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 600, marginBottom: 16 }}>
              Protocolo Padronizado de Registro de Atendimento — {AREAS[current.area]?.label}
            </div>
            <RecordReadView record={current} />
          </div>
        )}
        {mode === "report" && currentPatient && (
          <div style={{ padding: 24 }}>
            <div style={{ textAlign: "center", fontFamily: "'Fraunces', serif", fontSize: 18, fontWeight: 600, marginBottom: 16 }}>Relatório de Acompanhamento</div>
            <ReportView patientName={currentPatient.nome} sessions={currentPatientSessions} />
          </div>
        )}
      </div>

      {confirmDelete && (
        <div className="no-print fixed inset-0 flex items-center justify-center px-4" style={{ background: "rgba(29,42,40,0.45)", zIndex: 50 }}>
          <div style={{ background: "white", borderRadius: 12, padding: 22, maxWidth: 340 }}>
            <div style={{ fontFamily: "'Fraunces', serif", fontSize: 16, marginBottom: 6, color: T.ink }}>Excluir este atendimento?</div>
            <div style={{ fontSize: 13.5, color: T.inkSoft, marginBottom: 18 }}>Essa ação não pode ser desfeita.</div>
            <div className="flex justify-end gap-3">
              <button onClick={() => setConfirmDelete(null)} style={{ fontSize: 13.5, color: T.inkSoft }}>Cancelar</button>
              <button onClick={() => doDelete(confirmDelete)} style={{ fontSize: 13.5, color: "white", background: T.warn, borderRadius: 7, padding: "7px 14px" }}>Excluir</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
