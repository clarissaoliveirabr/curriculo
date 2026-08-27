const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  BorderStyle, LevelFormat, convertInchesToTwip,
} = require("docx");

const FONT = "Calibri";
const NAVY = "1F3A5F";
const GRAY = "444444";

const hr = () => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: NAVY, space: 4 } },
  spacing: { after: 160 },
});

const sectionHeading = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 260, after: 100 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY, space: 2 } },
  children: [new TextRun({ text, bold: true, color: NAVY, size: 22, font: FONT, allCaps: true })],
});

const jobHeader = (org, role, dates, note) => new Paragraph({
  spacing: { before: 160, after: 20 },
  tabStops: [{ type: "right", position: convertInchesToTwip(6.5) }],
  children: [
    new TextRun({ text: org.toUpperCase(), bold: true, size: 21, font: FONT }),
    new TextRun({ text: `\t${dates}`, bold: true, size: 19, font: FONT, color: GRAY }),
  ],
});

const roleLine = (role, extra) => new Paragraph({
  spacing: { after: 60 },
  children: [
    new TextRun({ text: role, italics: true, size: 20, font: FONT, color: GRAY }),
    ...(extra ? [new TextRun({ text: `  ${extra}`, size: 18, font: FONT, color: "B00020" })] : []),
  ],
});

const bullet = (text) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { after: 60 },
  children: [new TextRun({ text, size: 20, font: FONT })],
});

const body = (text) => new Paragraph({
  spacing: { after: 120 },
  children: [new TextRun({ text, size: 20, font: FONT })],
});

const compLine = (label, text) => new Paragraph({
  spacing: { after: 60 },
  children: [
    new TextRun({ text: `${label}: `, bold: true, size: 20, font: FONT }),
    new TextRun({ text, size: 20, font: FONT }),
  ],
});

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 260 } } } }],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 720, bottom: 720, left: 900, right: 900 },
      },
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [new TextRun({ text: "CLARISSA OLIVEIRA", bold: true, size: 34, font: FONT, color: NAVY })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 60 },
        children: [new TextRun({ text: "[TÍTULO-ALVO — ex.: Gerente de Produtos | Inovação | Negócios]", bold: true, size: 21, font: FONT, color: "B00020" })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 160 },
        children: [new TextRun({
          text: "Manaus, AM (disponível para mudança)  |  (92) 99245-9624  |  clarissaoliveira.br@gmail.com  |  linkedin.com/in/clarissaoliveirabr",
          size: 18, font: FONT, color: GRAY,
        })],
      }),
      hr(),

      sectionHeading("Resumo Profissional"),
      body("Executiva com 12+ anos de experiência em [gestão de produtos / desenvolvimento de negócios / inovação — ajuste conforme a vaga], especializada em conectar mercado, estratégia e times técnicos para gerar [resultado-alvo da vaga: crescimento de receita / lançamento de produtos / eficiência operacional]. Histórico em inteligência de mercado, precificação, go-to-market, Open Innovation e PD&I, com entregas mensuráveis em expansão comercial, estruturação de operações e desenvolvimento de portfólio."),

      sectionHeading("Competências"),
      compLine("Negócios e Comercial", "desenvolvimento B2B, inteligência de mercado, precificação, planejamento comercial, parcerias estratégicas"),
      compLine("Produtos e Tecnologia", "Product Discovery, levantamento de requisitos, MVP, go-to-market, produtos digitais, transformação digital"),
      compLine("Inovação e Projetos", "Open Innovation, PD&I, ecossistema de startups, metodologias ágeis (Scrum), OKR, gestão de stakeholders"),
      compLine("Idiomas", "[inserir nível — ex.: Inglês avançado / Espanhol intermediário]"),
      compLine("Ferramentas", "[inserir — ex.: Excel avançado, Power BI, Jira, Miro, CRM utilizado]"),

      sectionHeading("Experiência Profissional (Tempo Integral)"),

      jobHeader("ILAIN — Instituto Latino Americano de Inovação", "Diretora de Operações", "03/2026 – 08/2026"),
      roleLine("Diretora de Operações"),
      bullet("Estruturou a operação de uma ICT de tecnologia e PD&I do zero, definindo [X] processos e frentes de execução, o que permitiu [Y — ex.: reduzir em Z% o tempo de estruturação institucional]."),
      bullet("Definiu portfólio de serviços e posicionamento para [X] empresas-cliente, conectando demanda de mercado e viabilidade técnica, resultando em [Y]."),
      bullet("Conduziu a articulação entre mercado, academia e setor produtivo e apoiou o processo de credenciamento CAPDA, entregando [resultado/status do credenciamento]."),

      jobHeader("Innyx Tecnologia Ltda", "", "08/2024 – 03/2026"),
      roleLine("Gerente de Inovação e Produtos — Brasil  (08/2024 – 03/2026)"),
      bullet("Liderou o ciclo completo de Open Innovation e produtos — da identificação de oportunidades ao lançamento — entregando [X] iniciativas/produtos em [período]."),
      bullet("Mapeou necessidades de clientes e usuários e traduziu em requisitos, reduzindo [Y — ex.: tempo de descoberta] ao conectar mercado, estratégia e times técnicos."),
      bullet("Definiu prioridades e proposta de valor do portfólio com base em dados de clientes, tendências e viabilidade de negócio, direcionando [X]% do roadmap."),
      bullet("Acompanhou concorrência, métricas e ROI para embasar decisões de investimento em produto, apoiando [Y — resultado de negócio]."),
      roleLine("Coordenadora Comercial — Brasil  (06/2023 – 07/2024)"),
      bullet("Analisou performance comercial, custos, margens e concorrência, gerando planos de ação que resultaram em [X% de melhoria em margem/performance]."),
      bullet("Conduziu apresentações e treinamentos estratégicos entre Comercial, Marketing, Produto e Sourcing, alinhando [X] times em torno de [objetivo]."),

      jobHeader("Faculdade IEDi", "", "03/2024 – 08/2024"),
      roleLine("Coordenadora Comercial — Região Norte", "⚠ confirmar datas — ver observação no e-mail/relatório"),
      bullet("Estruturou a estratégia comercial dos polos de Manaus e Boa Vista, aumentando em 42% o número de alunos em 4 meses."),
      bullet("Conduziu inteligência de mercado e ações de captação, acompanhando a operação comercial regional."),

      jobHeader("Itriad Pesquisa e Desenvolvimento", "", "01/2022 – 06/2023"),
      roleLine("Analista de Negócios Pleno"),
      bullet("Planejou e acompanhou atividades, riscos e documentação de projetos, garantindo alinhamento entre negócio e desenvolvimento em [X] projetos simultâneos."),
      bullet("Atuou como ponte entre negócio e times técnicos, detalhando requisitos que reduziram [Y — ex.: retrabalho/tempo de entrega]."),

      jobHeader("SEMTRAD/SEMTEPI — Prefeitura de Manaus", "", "04/2019 – 08/2020"),
      roleLine("Líder de Eixo de Inovação"),
      bullet("Conduziu programas de inovação aberta, mentorias e editais para empreendedores e startups, articulando poder público, empreendedores e instituições."),
      bullet("Criou a metodologia IMPULSO, aplicada a mais de 3.500 jovens e 45 startups."),

      jobHeader("Instituto CERTI Amazônia", "", "08/2015 – 08/2019"),
      roleLine("Analista de Projetos — Negócios e Empreendedorismo"),
      bullet("Desenvolveu programas de empreendedorismo e projetos de P&D, incluindo mentorias para [X] startups e iniciativas de inovação corporativa."),
      bullet("Conectou demandas de mercado a estruturação e validação de novos negócios, apoiando [X] validações/lançamentos."),

      sectionHeading("Atuação Paralela — Empreendedorismo e Voluntariado"),
      body("Separado da experiência em tempo integral para deixar claro o regime de cada vínculo — evita a leitura de vínculos empregatícios sobrepostos."),

      jobHeader("Hangar Brasil", "", "03/2022 – atual"),
      roleLine("Fundadora e Consultora em Negócios, Produtos e Inovação — atuação autônoma/PJ, paralela"),
      bullet("Conduz diagnósticos estratégicos para clientes, transformando necessidades em propostas de valor, produtos e planos de execução, com [X] clientes atendidos."),
      bullet("Estruturou posicionamento e go-to-market para [X] empresas/startups, apoiando [Y — resultado de captação/crescimento]."),
      bullet("Idealizou “O Grande Pitch”, programa de inovação exibido pelo SBT Amazonas."),

      jobHeader("ABRH Amazonas", "", "12/2024 – 06/2026"),
      roleLine("Diretora de Projetos e Inovação — atuação voluntária"),
      bullet("Articulou lideranças e parceiros para planejamento e execução de projetos institucionais de inovação e desenvolvimento de negócios."),

      sectionHeading("Projetos e Resultados Selecionados"),
      bullet("O Grande Pitch — idealização e estruturação de reality de empreendedorismo, conectando startups a desafios de mercado e oportunidades comerciais; exibido pelo SBT Amazonas."),
      bullet("Metodologia IMPULSO — metodologia própria de desenvolvimento de negócios inovadores, aplicada a 3.500+ jovens e 45 startups."),
      bullet("Programas de Open Innovation — desenho e execução de conexões entre empresas, startups e instituições para validação de soluções."),

      sectionHeading("Formação e Certificações"),
      body("Engenharia de Produção — Estácio | em andamento, previsão de conclusão em 2028"),
      body("Certificações: Triple Track One Team (2024); Team OKR (2024); Lean Inception CLF — Caroli (2024); Scrum Master (2019); Green Belt Six Sigma (2014)"),

      sectionHeading("Metodologias e Ferramentas"),
      body("Design Thinking · Lean Startup · Lean Inception · Product Discovery · MVP · Jornada do Cliente · Proposta de Valor · Modelagem de Negócios · Scrum · OKR · Gestão Ágil · Six Sigma"),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 160 },
        children: [new TextRun({ text: "Disponibilidade para viagens e mudança de endereço", italics: true, size: 18, font: FONT, color: GRAY })],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync(process.argv[2] || "output.docx", buf);
  console.log("done");
});
