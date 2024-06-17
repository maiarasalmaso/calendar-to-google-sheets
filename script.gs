function extrairEventosConcluidosAgosto() {
  var calendarioId='INSIRA SEU ID AQUI';
  var calendario=CalendarApp.getCalendarById(calendarioId);

  if(!calendario){
    SpreadsheetApp.getActiveSpreadsheet().toast('O calendário não foi encontrado. Verifique o ID do calendário.', 'Erro', 5);
    return;
  }

  var dataInicioAgosto=new Date('2024-08-01');
  var dataFimAgosto=new Date('2024-08-31');

  var eventosAgosto=calendario.getEvents(dataInicioAgosto,dataFimAgosto);

  var planilhaAgosto=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('agosto');
  if(!planilhaAgosto){
    planilhaAgosto=SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    planilhaAgosto.setName('agosto');
  }else{
    planilhaAgosto.clear();
  }

  var cabecalhos=['Sala Reservada','Reservado por','Paciente','Valor da Terapia','Data Reservada','Hora Reservada'];
  planilhaAgosto.appendRow(cabecalhos);

  var somaPorReservadoPor={}; // Objeto para armazenar a soma dos valores por "Reservado por"

  eventosAgosto.forEach(function(evento){
    var descricao=evento.getDescription();
    var reservadoPor='';
    var paciente='';
    var valorTerapia=0; // Inicializa o valor da terapia como zero

    var reservadoPorMatch=descricao.match(/<b>Reservado por<\/b>\n([^<\n]+)/);
    if(reservadoPorMatch&&reservadoPorMatch.length>1){
      reservadoPor=reservadoPorMatch[1].replace(/<br>/g,'');
    }

    var pacienteMatch=descricao.match(/<b>Paciente<\/b>\n([^\n]+)/);
    if(pacienteMatch&&pacienteMatch.length>1){
      paciente=pacienteMatch[1];
    }

    var valorTerapiaMatch=descricao.match(/<b>Valor da terapia<\/b>\n([^\n]+)/);
    if(valorTerapiaMatch&&valorTerapiaMatch.length>1){
      valorTerapia=parseFloat(valorTerapiaMatch[1]); // Converte o valor para número
    }

    if(!somaPorReservadoPor[reservadoPor]){
      somaPorReservadoPor[reservadoPor]=0;
    }

    somaPorReservadoPor[reservadoPor]+=valorTerapia; // Adiciona o valor à soma

    var titulo=evento.getTitle();
    var sala='';

    var salaMatch=titulo.match(/Sala \d+/);
    if(salaMatch&&salaMatch.length>0){
      sala=salaMatch[0];
    }

    var linha=[
      sala,
      reservadoPor,
      paciente,
      valorTerapia,
      evento.getStartTime().toLocaleDateString(),
      evento.getStartTime().toLocaleTimeString()
    ];
    planilhaAgosto.appendRow(linha);
  });

  Object.keys(somaPorReservadoPor).forEach(function(reservadoPor){
    var valorTotal=somaPorReservadoPor[reservadoPor];
    var valorDesconto=valorTotal*0.30; // Calcula 30% do valor total
    planilhaAgosto.appendRow(['',reservadoPor,'',valorTotal,'','']); // Adiciona a soma
    planilhaAgosto.appendRow(['','30% de '+reservadoPor,'',valorDesconto,'','']); 
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('Eventos concluídos para agosto extraídos com sucesso!', 'Concluído', 5);
}
