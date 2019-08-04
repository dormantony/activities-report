
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generar reporte semana')
  .addItem('Semana actual', 'semanaActual')  
  .addItem('Elegir semana', 'elegirSemana')
  .addToUi();
}
Date.prototype.getWeek = function() {
    var onejan = new Date(this.getFullYear(),0,1);
    return Math.ceil((((this - onejan) / 86400000) + onejan.getDay()+1)/7);
} 
function semanaActual(){
  now=new Date()
  generarReporte(now.getWeek())
}
function elegirSemana(){
//  var NUMSEMANA=Browser.inputBox("Introduce el numero de semana");
generarReporte(31)
  
}
function prinbt(){
 console.log("probando") 
}
function generarReporte(NUMSEMANA){
   console.log("probando")  
  SpreadsheetApp.getActive().getSheetByName("Reporte").getRange("A7").setValue(NUMSEMANA)
  sheettemas=SpreadsheetApp.getActive().getSheetByName("Temas trabajando");
  lastrow=sheettemas.getLastRow();
  lastColumn=sheettemas.getLastColumn();
  temp=SpreadsheetApp.getActive().getSheetByName("temp");
  rangetemas=sheettemas.getRange("b1:b"+lastrow);
  indexes=sheettemas.getRange(3, 1, 2, lastColumn);
//  console.log("sucediendo bien2"+indexes.getValues()[0].length)
//  console.log("iteraciones"+indexes.getValues()[0].length);
  inicio=0;
  fin=0;
  inicioEncontrado=false;
  elementos=0;
  contador=0;
  indiceMes=0;

//  Se busca el rango de la semana para pegarse
  for(i=0; i<indexes.getValues()[0].length;i++){
    if(indexes.getValues()[0][i]==NUMSEMANA){
     elementos++;
     console.log("SEMANA ENCONTRADA EN"+i)
     if(inicioEncontrado==false){
       inicioEncontrado=true;
       inicio=i;
     }else{
     fin=i;
     }
    }else if(indexes.getValues()[0][i]==""){
      console.log("indiceMes:"+i);
      indiceMes=i;
    }
    
    if(elementos==7){
       console.log("se encontraron todos los elementos")
     break;
     }
  }
  numRows=rangetemas.getValues().length;
  inicio++;
  fin++;
  console.log("inicio:"+inicio+" fin:"+fin+" numRows"+numRows+" elementos"+elementos);
  temp.getRange(1, 1, numRows).setValues(rangetemas.getValues());
  //  Una vez con el rango se copia y pega en la hoja
  borrarColumna=(indiceMes-inicio)+1
  rangosemana=sheettemas.getRange(1, inicio, sheettemas.getLastRow(), elementos)
  temp.getRange(1, 2, numRows,elementos ).setValues(rangosemana.getValues())
  console.log(borrarColumna)
  
  Browser.msgBox("Reporte generado de semana:"+NUMSEMANA);
}