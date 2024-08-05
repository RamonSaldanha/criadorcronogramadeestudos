const xlsx = require('xlsx');


const texto = `DIREITO CONSTITUCIONAL: Constituição: conceito, objeto e classificações; supremacia da Constituição; aplicabilidade, vigência e eficácia das normas constitucionais; interpretação constitucional. Princípios fundamentais. Ações Constitucionais: habeas corpus, habeas data, mandado de segurança; mandado de injunção; ação popular; ação civil pública. Controle de constitucionalidade: sistemas difuso e concentrado; ação direta de inconstitucionalidade; ação declaratória de constitucionalidade; arguição de descumprimento de preceito fundamental; súmula vinculante; repercussão geral. Direitos e garantias fundamentais: direitos e deveres individuais e coletivos; direitos sociais; direitos de nacionalidade; direitos políticos; partidos políticos. Organização político- administrativa: União; Estados; Municípios; Distrito Federal; Territórios; intervenção federal e estadual. Administração Pública: disposições gerais; servidores públicos. Organização dos Poderes. Poder Executivo: atribuições e responsabilidades do Presidente da República. Poder Legislativo: órgãos e atribuições; processo legislativo; fiscalização contábil, financeira e orçamentária. Poder Judiciário: disposições gerais; Supremo Tribunal Federal; Conselho Nacional de Justiça; Superior Tribunal de Justiça; Tribunais Regionais Federais e Juízes Federais; Tribunais e Juízes Eleitorais; Tribunais e Juízes dos Estados. Funções essenciais à Justiça: Ministério Público; Advocacia Pública; Advocacia; Defensoria Pública. Finanças Públicas: normas gerais; dos orçamentos. Ordem econômica e financeira: princípios gerais da atividade econômica; política urbana; política agrícola e fundiária e reforma agrária. Ordem social: disposição geral; seguridade social; educação, cultura e desporto; comunicação social; meio ambiente; indígenas. DIREITO ADMINISTRATIVO: Administração pública: princípios básicos. Poderes administrativos: poder hierárquico, poder disciplinar, poder regulamentar, poder de polícia, uso e abuso do poder. Serviços públicos: conceito, regime jurídico, princípios, titularidade e competência. Delegação: concessão, permissão e autorização. Ato administrativo: conceito, requisitos e atributos; anulação, revogação e convalidação; discricionariedade e vinculação. Organização administrativa: administração direta e indireta; centralizada e descentralizada; autarquias, fundações, empresas públicas, sociedades de economia mista, consórcios públicos (Lei nº 11.107/2005). Órgãos públicos: conceito, natureza e classificação. Servidores públicos: cargo, emprego e função públicos. Lei nº 8.112/1990 (Regime Jurídico dos Servidores Públicos Civis da União e alterações): disposições preliminares, provimento, vacância, remoção, redistribuição e substituição; direitos e vantagens: vencimento e remuneração; vantagens; férias; licenças; afastamentos; direito de petição; regime disciplinar: deveres e proibições; acumulação; responsabilidades; penalidades. Processo administrativo (Lei nº 9.784/1999): disposições gerais, direitos e deveres dos administrados. Controle e responsabilização da administração: controle administrativo; controle judicial; controle legislativo; responsabilidade civil do Estado. Lei nº 8.429/1992: disposições gerais; atos de improbidade administrativa. Lei nº 11.416/2006, que dispõe sobre as carreiras do Poder Judiciário da União. Licitações e Contratos da Administração Pública – Lei nº 14.133/2021 e suas alterações. Convênios administrativos. Pregão: Lei n° 10.520/2002. Regime Diferenciado de Contratações Públicas: Lei Federal n 12.462, de 4 de agosto de 2011. Parcerias Público-Privadas (Lei nº 11.079/2004, com alterações posteriores). Bens públicos: regime jurídico; classificação; administração; aquisição e alienação; utilização; autorização de uso, permissão de uso, concessão de uso, concessão de direito real de uso e cessão de uso. Intervenção do Estado na propriedade: desapropriação; servidão administrativa; tombamento; requisição administrativa; ocupação temporária; limitação administrativa. Terceiro Setor: Organizações Sociais (Lei nº 9.637/1998). Organizações da Sociedade Civil de Interesse Público (Lei nº 9.790/1999, com alterações posteriores). Parcerias entre a administração pública e as organizações da sociedade civil: Lei 13.019/2014. Mandado de Segurança individual. Mandado de Segurança Coletivo. Ação Popular. Ação Civil Pública. Mandado de Injunção. Habeas Data. DIREITO CIVIL: Lei. Eficácia da lei. Aplicação da lei no tempo e no espaço. Interpretação da lei. Lei de Introdução às normas do Direito Brasileiro. Das Pessoas Naturais: Da Personalidade e Da Capacidade. Dos Direitos da Personalidade. Das pessoas jurídicas. Domicílio Civil. Bens. Dos Fatos Jurídicos: Dos negócios jurídicos; Dos atos jurídicos lícitos. Dos Atos Ilícitos. Prescrição e decadência. Do Direito das Obrigações. Dos Contratos: Das Disposições Gerais; Da Compra e Venda; Da Prestação de Serviço; Do Mandato; Da Transação. Empreitada (cap. VIII do Título VI do CC). Da Responsabilidade Civil. Do Penhor, Da Hipoteca e Da Anticrese. DIREITO PROCESSUAL CIVIL: Novo Código de Processo Civil - Lei Federal n° 13.105/2015 e alterações e legislações especiais. Princípios gerais do processo civil. Fontes. Lei processual civil. Eficácia. Aplicação. Interpretação. Direito Processual Intertemporal. Critérios. Jurisdição. Conceito. Característica. Natureza jurídica. Princípios. Limites. Competência. Critérios determinadores. Competência originária dos Tribunais Superiores. Competência absoluta e relativa. Modificações. Meios de declaração de incompetência. Conflitos de competência e de atribuições. Direito de ação. Elementos. Condições. Classificação e critérios identificadores. Concurso e cumulação de ações. Conexão e continência. Processo: Noções gerais. Relação Jurídica Processual. Pressupostos Processuais. Processo e procedimento. Espécies de processos e de procedimentos. Objeto do processo. Mérito. Questão principal, questões preliminares e prejudiciais. Sujeitos Processuais. Juiz. Mediadores e Conciliadores. Princípios. Poderes. Deveres. Responsabilidades. Impedimentos e Suspeição. Organização judiciária federal e estadual. Sujeitos Processuais. Partes e Procuradores. Capacidade e Legitimação. Representação e Substituição Processual. Litisconsórcio. Da Intervenção de Terceiros. Da Assistência. Da Denunciação da Lide. Do Chamamento ao Processo. Do incidente de desconsideração da personalidade jurídica. Do Amicus Curiae. Advogado. Ministério Público. Auxiliares da Justiça. A Advocacia Pública. Prerrogativas da Fazenda Pública em juízo. Fatos e atos processuais. Forma. Tempo. Lugar. Prazos. Comunicações. Nulidades. Procedimento comum. Aspectos Gerais. Fases. Petição inicial. Requisitos. Indeferimento da petição inicial e improcedência liminar do pedido. Resposta do réu. Impulso processual. Prazos e preclusão. Prescrição. Inércia processual: contumácia e revelia. Formação, suspensão e extinção do processo. Contestação. Reconvenção. Das Providências preliminares e do Saneamento. Julgamento conforme o estado do processo. Provas. Audiências. Conciliação e Mediação. Instrução e julgamento. Distribuição do ônus da prova. Fatos que independem de prova. Depoimento pessoal. Confissão. Prova documental. Exibição de documentos ou coisas. Prova testemunhal. Prova pericial. Inspeção judicial. Exame e valoração da prova. Produção Antecipada de Provas. Da Tutela Provisória: Tutelas de Urgência e de Evidência. Fungibilidade. Princípios Gerais. Protesto, notificação e interpelação. Arresto. Sequestro. Caução. Busca e Apreensão. Exibição. Justificação. Sentença. Conceito. Classificações. Requisitos. Efeitos. Publicação, intimação, correção e integração da sentença. Do cumprimento da Sentença. Coisa julgada. Conceito. Espécies. Limites. Remessa Necessária. Meios de impugnação à sentença. Ação rescisória. Recursos. Disposições Gerais. Apelação. Agravos. Embargos de Declaração. Embargos de Divergência. Recurso Ordinário. Recurso Especial. Recurso Extraordinário. Recursos nos Tribunais Superiores. Reclamação e correição. Repercussão geral. Súmula vinculante. Recursos repetitivos. Liquidação de Sentença. Espécies. Procedimento. Cumprimento da sentença. Procedimento. Impugnação. Processo de Execução. Princípios gerais. Espécies. Execução contra a Fazenda Pública. Regime de Precatórios. Requisições de Pequeno Valor. Execução de obrigação de fazer e de não fazer. Execução por quantia certa. Embargos de Terceiros. Exceção de pré- executividade. Remição. Suspensão e extinção do processo de execução. Procedimentos Especiais. Generalidades. Características. Espécies. Ação de Consignação em Pagamento. Ação Monitória. Ação de Exigir Contas. Ações Possessórias. Restauração de autos. Ação Popular. Ação Civil Pública. Aspectos processuais. Mandado de Segurança. Mandado de Injunção. Mandado de Segurança Coletivo. Habeas Data. O Processo Civil nos sistemas de controle da constitucionalidade. Ação Direta de Inconstitucionalidade. Ação Declaratória de Constitucionalidade. Medida Cautelar. Declaração incidental de inconstitucionalidade. Ações Civis Constitucionais. Arguição de Descumprimento de Preceito Fundamental. Ação de Improbidade Administrativa. Jurisprudência dominante dos Tribunais Superiores em matéria de Processo Civil aplicáveis ao novo código de Processual Civil e demais procedimentos previstos em legislação processual específica. DIREITO PENAL: Da aplicação da lei penal. Do Crime. Da imputabilidade penal. Do concurso de pessoas. Das Penas: Das espécies de pena; Da cominação das penas; Dos efeitos da condenação. Da Ação penal. Da extinção da punibilidade. Dos crimes contra a fé pública: Da falsidade documental. Dos crimes contra a Administração Pública: Dos crimes praticados por funcionário público contra a administração em geral; Dos crimes praticados por particular contra a administração em geral; Dos crimes contra a administração da Justiça. Abuso de autoridade (Lei nº 4.898/1965 e alterações posteriores). Direito Processual Penal: Princípios gerais: aplicação da lei processual no tempo, no espaço em relação às pessoas; sujeitos da relação processual. Do Inquérito policial. Da ação penal. Da competência. Da prova: Do exame de corpo de delito e das perícias em geral; Do interrogatório do acusado; Das testemunhas; Dos documentos; Da busca e da apreensão. Do Juiz, do Ministério Público, Do acusado e defensor, dos Assistentes e Auxiliares da Justiça. Da prisão e da liberdade provisória. Das citações e intimações. Da sentença. Das nulidades. Dos recursos em geral: disposições gerais; do recurso em sentido estrito; da apelação; do habeas corpus e seu processo. Dos Juizados Especiais Criminais (Lei nº 9.099/1995 e alterações posteriores e Lei nº 10.259/2001 e alterações posteriores). Súmulas do STJ e do STF. DIREITO PREVIDENCIÁRIO: Seguridade social: origem e evolução legislativa no Brasil; conceito; organização e princípios constitucionais. Da assistência social. Dos regimes de previdência social existentes. Regime Geral da Previdência Social: beneficiário, benefícios em espécie e custeio (Leis nº 8.212/91 e 8.213/91). Seguridade Social do Servidor Público: noções gerais, benefícios e custeio. Previdência Complementar (Lei Complementar nº 109/2001). Relação entre a União, os Estados, o Distrito Federal e os Municípios, suas autarquias, fundações, sociedades de economia mista e outras entidades públicas e suas respectivas entidades fechadas de previdência complementar (Lei Complementar nº 108/2001). Lei nº 12.618/2012 (Regime de Previdência Complementar para os Servidores Públicos Federais). Impactos da Lei nº 13.467/2017 na Previdência Social. DIREITO TRIBUTÁRIO: Normas gerais de Direito Tributário. Fontes do Direito Tributário. Norma tributária: vigência, aplicação, interpretação e integração. Tributo: conceito, natureza jurídica e espécies. Hipótese de incidência: conceito e aspectos. Fato gerador. Obrigações tributárias: conceito e espécies, sujeitos ativo e passivo. Obrigação principal e acessória. Crédito tributário: conceito, natureza, lançamento, modalidades e revisão do lançamento, suspensão, extinção e exclusão do crédito tributário. Responsabilidade tributária. O Sistema Tributário Nacional: limitações constitucionais ao poder de tributar, imunidade tributária, competência tributária, tributos federais, estaduais e municipais. Administração tributária. Repartição das receitas tributárias. Garantias e privilégios do crédito tributário. DIREITO DO CONSUMIDOR: Do Código de Defesa do Consumidor. Dos direitos do consumidor. Das disposições gerais. Dos direitos básicos do consumidor. Da qualidade de produtos e serviços. Da preservação e da reparação de danos (da proteção à saúde e segurança). Da responsabilidade pelo fato do produto e do serviço. Da responsabilidade por vício do produto e do serviço. Da decadência e da prescrição. Da desconsideração da personalidade jurídica. Das práticas comerciais (das disposições gerais). Da oferta. Da publicidade. Das práticas abusivas. Da cobrança de dívidas. Da proteção contratual: disposições gerais. Das cláusulas abusivas. Dos contratos de adesão. Da defesa do consumidor em juízo. Das disposições do Código de Defesa do Consumidor relacionadas à defesa do consumidor em juízo. Das ações coletivas para a defesa de interesses individuais homogêneos. Das ações de responsabilidade do fornecedor de produtos e serviços. Da tutela específica nas obrigações de fazer ou não fazer. Da sentença. Da coisa julgada. Da liquidação da sentença coletiva. Do cumprimento da sentença. Noção de verossimilhança e hipossuficiência para facilitação da defesa em juízo dos direitos do consumidor, inclusive com a inversão do ônus da prova. Sanções administrativas e penais: Da competência concorrente. Multa, apreensão, inutilização, cassação de registro, proibição de fabricação, suspensão temporária de atividade, revogação ou cassação de concessão ou permissão, da interdição. Da Contrapropaganda. O sistema nacional de defesa do consumidor: A política nacional de relações de consumo - SNDC e PROCON.`;

function divideDisciplinasAssuntos(texto) {
  const regexDisciplina = /([A-ZÁÉÍÓÚÂÊÎÔÛÀÈÌÒÙÃÕÇÑ\s-]+:)/g;
  const disciplinas = texto.split(regexDisciplina).filter(Boolean);

  let estruturaFinal = [];

  for (let i = 0; i < disciplinas.length; i += 2) {
    let disciplina = disciplinas[i].trim();
    let assuntos = disciplinas[i + 1].trim().split(/\. /).filter(Boolean).map(assunto => assunto + '.');
    estruturaFinal.push({ disciplina, assuntos });
  }

  return estruturaFinal;
}

function criarCronograma(dadosDisciplinas) {
  const semanas = 11;
  const diasPorSemana = 5;
  const totalDias = semanas * diasPorSemana;
  let cronograma = new Array(totalDias).fill(null).map((_, index) => ({
      semana: Math.floor(index / diasPorSemana) + 1,
      dia: ['seg', 'ter', 'qua', 'qui', 'sex'][index % diasPorSemana],
      disciplina: '',
      assuntos: []
  }));

  let disciplinasRotativas = dadosDisciplinas.slice(); // Cópia para manipulação
  let indiceDia = 0;
  let ultimoIndiceUsado = -1;

  while (indiceDia < totalDias) {
      if (ultimoIndiceUsado + 1 >= disciplinasRotativas.length) {
          // Reorganizar e resetar o índice se chegou ao fim da lista
          disciplinasRotativas = disciplinasRotativas.sort(() => 0.5 - Math.random());
          ultimoIndiceUsado = -1;
      }
      
      let disciplinaAtual = disciplinasRotativas[ultimoIndiceUsado + 1];
      let maxAssuntosPorDia = Math.min(10, disciplinaAtual.assuntos.length); // Determina um limite máximo baseado na disponibilidade
      
      cronograma[indiceDia].disciplina = disciplinaAtual.disciplina;
      cronograma[indiceDia].assuntos = disciplinaAtual.assuntos.splice(0, maxAssuntosPorDia);

      ultimoIndiceUsado++;
      indiceDia++;

      // Evitar repetição consecutiva verificando o próximo dia
      if (indiceDia < totalDias && disciplinasRotativas[ultimoIndiceUsado % disciplinasRotativas.length].disciplina === cronograma[indiceDia - 1].disciplina) {
          ultimoIndiceUsado++; // Pular para a próxima disciplina para o próximo dia
      }
  }

  // Checar se há dias sem assuntos e tentar redistribuir de dias com muitos assuntos
  cronograma.forEach((dia, index) => {
      if (dia.assuntos.length === 0 && index > 0) {
          // Tenta pegar mais assuntos do dia anterior ou posterior
          let prevIndex = index - 1;
          let nextIndex = index + 1 < cronograma.length ? index + 1 : index;
          let sourceDay = cronograma[prevIndex].assuntos.length > 5 ? cronograma[prevIndex] : cronograma[nextIndex];

          if (sourceDay.assuntos.length > 5) {
              dia.assuntos = sourceDay.assuntos.splice(5); // Distribuir excesso
              dia.disciplina = sourceDay.disciplina;
          }
      }
  });

  return cronograma;
}

// Função para verificar se todos os assuntos estão no cronograma
function verificarCronograma(dadosDisciplinas, cronograma) {
  const assuntosOriginais = new Set(dadosDisciplinas.flatMap(d => d.assuntos.map(a => `${d.disciplina} ${a}`)));
  const assuntosNoCronograma = new Set(cronograma.flatMap(dia => dia.assuntos.map(a => `${dia.disciplina} ${a}`)));

  const faltando = Array.from(assuntosOriginais).filter(assunto => !assuntosNoCronograma.has(assunto));
  return faltando.length === 0 ? 'Todos os assuntos estão incluídos.' : `Assuntos faltando: ${faltando.join(', ')}`;
}
function assuntosFaltantesPorDisciplina(dadosDisciplinas, cronograma) {
  const assuntosPorDisciplina = new Map(dadosDisciplinas.map(d => [d.disciplina, new Set(d.assuntos)]));
  cronograma.forEach(dia => {
      dia.assuntos.forEach(assunto => {
          if (assuntosPorDisciplina.has(dia.disciplina)) {
              assuntosPorDisciplina.get(dia.disciplina).delete(assunto);
          }
      });
  });

  let resultado = [];
  assuntosPorDisciplina.forEach((assuntos, disciplina) => {
      if (assuntos.size > 0) {
          resultado.push({
              disciplina,
              assuntos: Array.from(assuntos)
          });
      }
  });
  return resultado;
}

// Atualizar a chamada de verificação no fim para incluir a nova função
const dadosDisciplinas = divideDisciplinasAssuntos(texto);
const cronograma = criarCronograma(dadosDisciplinas);
const assuntosFaltantes = assuntosFaltantesPorDisciplina(dadosDisciplinas, cronograma);

console.log(cronograma);
console.log(assuntosFaltantes);


function exportarParaExcel(dados, nomeArquivo) {
  // Transforma cada dia em múltiplas linhas, uma para cada assunto
  let rows = [];
  dados.forEach(dia => {
      if (dia.assuntos.length > 0) {
          // Adiciona a primeira linha com todas as informações
          rows.push({
              semana: dia.semana,
              dia: dia.dia,
              disciplina: dia.disciplina,
              assunto: dia.assuntos[0]
          });
          // Adiciona as linhas subsequentes com apenas o assunto
          dia.assuntos.slice(1).forEach(assunto => {
              rows.push({
                  semana: "",  // Vazio para semana, dia e disciplina nas linhas subsequentes
                  dia: "",
                  disciplina: "",
                  assunto
              });
          });
      }
  });

  // Cria a planilha e o livro a partir das linhas
  const ws = xlsx.utils.json_to_sheet(rows, {
      header: ["semana", "dia", "disciplina", "assunto"],  // Define os cabeçalhos para garantir ordem
      skipHeader: true  // Pula a linha de cabeçalho para uso interno
  });
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Dados");
  xlsx.writeFile(wb, `${nomeArquivo}.xlsx`);
}

// Exemplo de uso para exportar cronograma e assuntos faltantes
exportarParaExcel(cronograma, 'Cronograma');

exportarParaExcel(assuntosFaltantes.map(item => ({
  semana: '',
  dia: '',
  disciplina: item.disciplina,
  assuntos: item.assuntos
})), 'Assuntos_Faltantes');