function extrairEventosConcluidosOutubro() {
  var calendarioId = 'INSIRA SEU ID AQUI';
  var calendario = CalendarApp.getCalendarById(calendarioId);

  if (!calendario) {
    SpreadsheetApp.getActiveSpreadsheet().toast('O calendário não foi encontrado. Verifique o ID do calendário.', 'Erro', 5);
    return;
  }

  var dataInicioJunho = new Date('2024-08-01');
  var dataFimJunho = new Date('2024-08-31');

  var eventosJunho = calendario.getEvents(dataInicioJunho, dataFimJunho);

  // Ordenar eventos por data crescente
  eventosJunho.sort(function(a, b) {
    return a.getStartTime() - b.getStartTime();
  });

  var planilhaJunho = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('junho');
  if (!planilhaJunho) {
    planilhaJunho = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    planilhaJunho.setName('agosto');
  } else {
    planilhaJunho.clear();
  }

  var cabecalhos = ['Nome', 'Data', 'Hora', 'Paciente', 'Valor da Terapia'];
  planilhaJunho.appendRow(cabecalhos);
  var cabecalhoRange = planilhaJunho.getRange(1, 1, 1, cabecalhos.length);
  cabecalhoRange.setFontWeight('bold').setHorizontalAlignment('center');

  var eventosPorReservadoPor = {};

  // Agrupar eventos por "Reservado por"
  eventosJunho.forEach(function(evento) {
    var descricao = evento.getDescription();
    var reservadoPor = '';
    var paciente = '';
    var valorTerapia = 0; // Inicializa o valor da terapia como zero

    var reservadoPorMatch = descricao.match(/<b>Reservado por<\/b>\n([^<\n]+)/);
    if (reservadoPorMatch && reservadoPorMatch.length > 1) {
      reservadoPor = reservadoPorMatch[1].replace(/<br>/g, '');
    }

    var pacienteMatch = descricao.match(/<b>Paciente<\/b>\n([^\n]+)/);
    if (pacienteMatch && pacienteMatch.length > 1) {
      paciente = pacienteMatch[1];
    }

    var valorTerapiaMatch = descricao.match(/<b>Valor da terapia<\/b>\n([^\n]+)/);
    if (valorTerapiaMatch && valorTerapiaMatch.length > 1) {
      valorTerapia = parseFloat(valorTerapiaMatch[1]); // Converte o valor para número
    }

    if (!eventosPorReservadoPor[reservadoPor]) {
      eventosPorReservadoPor[reservadoPor] = [];
    }

    eventosPorReservadoPor[reservadoPor].push([
      reservadoPor,
      Utilities.formatDate(evento.getStartTime(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      evento.getStartTime().toLocaleTimeString(),
      paciente,
      valorTerapia
    ]);
  });

  // Inserir os eventos na planilha, agrupados por "Reservado por"
  var linhaAtual = 2;
  Object.keys(eventosPorReservadoPor).forEach(function(reservadoPor) {
    var linhasDoGrupo = eventosPorReservadoPor[reservadoPor].length;

    // Cabeçalho do grupo mesclado com a quantidade de reservas
    var rangeCabecalho = planilhaJunho.getRange(linhaAtual, 1, linhasDoGrupo, 1);
    rangeCabecalho.merge().setValue(reservadoPor + ' (' + linhasDoGrupo + ' reservas)').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

    // Adicionar os eventos do grupo
    eventosPorReservadoPor[reservadoPor].forEach(function(linha, index) {
      var rangeDados = planilhaJunho.getRange(linhaAtual + index, 2, 1, 4);
      rangeDados.setValues([linha.slice(1)]).setHorizontalAlignment('center');
    });

    // Aplicar bordas externas ao grupo de eventos
    var rangeGrupo = planilhaJunho.getRange(linhaAtual, 1, linhasDoGrupo, 5);
    rangeGrupo.setBorder(true, true, true, true, true, true);

    // Atualizar a linha atual para o próximo grupo
    linhaAtual += linhasDoGrupo + 1;

    // Inserir uma linha em branco abaixo de cada grupo
    if (linhasDoGrupo > 0) {
      planilhaJunho.insertRowAfter(linhaAtual - 1);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Eventos concluídos para junho extraídos com sucesso!', 'Concluído', 5);
}


