function extrairEventosConcluidosDezembro() {
  var calendarioId = 'INSIRA SEU ID AQUI';
  var calendario = CalendarApp.getCalendarById(calendarioId);

  var eventos = calendario.getEvents(new Date('2024-11-30'), new Date('2024-12-31')).sort((a, b) => a.getStartTime() - b.getStartTime());

  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('dezembro') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('dezembro');
  planilha.clear();

  planilha.appendRow(['Nome', 'Data', 'Hora', 'Paciente', 'Valor da Terapia']).getRange(1, 1, 1, 5)
    .setFontWeight('bold').setHorizontalAlignment('center');

  var eventosPorReservadoPor = {};

  eventos.forEach(evento => {
    var descricao = evento.getDescription();
    var reservadoPor = (descricao.match(/<b>Reservado por<\/b>\n([^<\n]+)/) || [])[1] || '';
    var paciente = (descricao.match(/<b>Paciente<\/b>\n([^\n]+)/) || [])[1] || '';
    var valorTerapia = parseFloat((descricao.match(/<b>Valor da terapia<\/b>\n([^\n]+)/) || [])[1]) || 0;

    if (evento.getGuestList().every(convidado => convidado.getGuestStatus() === CalendarApp.GuestStatus.NO)) return;

    if (!eventosPorReservadoPor[reservadoPor]) eventosPorReservadoPor[reservadoPor] = [];
    eventosPorReservadoPor[reservadoPor].push([
      reservadoPor,
      Utilities.formatDate(evento.getStartTime(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      evento.getStartTime().toLocaleTimeString(),
      paciente,
      valorTerapia
    ]);
  });

  var linhaAtual = 2;
  Object.keys(eventosPorReservadoPor).forEach(reservadoPor => {
    var grupo = eventosPorReservadoPor[reservadoPor];
    planilha.getRange(linhaAtual, 1, grupo.length, 1).merge().setValue(`${reservadoPor} (${grupo.length} reservas)`)
      .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

    grupo.forEach((linha, index) => {
      planilha.getRange(linhaAtual + index, 2, 1, 4).setValues([linha.slice(1)]).setHorizontalAlignment('center');
    });

    planilha.getRange(linhaAtual, 1, grupo.length, 5).setBorder(true, true, true, true, true, true);
    linhaAtual += grupo.length + 1;
    if (grupo.length > 0) planilha.insertRowAfter(linhaAtual - 1);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Eventos concluídos para dezembro extraídos com sucesso!', 'Concluído', 5);
}
