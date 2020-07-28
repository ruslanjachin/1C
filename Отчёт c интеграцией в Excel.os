
Перем ДатаНачала, ДатаКонца;  

//*******************************************
Процедура ВыгрузкаВЭксель() 
	
	// Константы 
	True = 1;
	False = 0;
	
	// Экселевские константы  
	xlContinuous = 1; 	 // Сплошная линия 
	xlToRight = -4161;   	 // Смещение вправо        
	xlHAlignCenter = -4108;  // Горизонтальное выравнивание по центру
	xlAscending = 1;
	xlYes = 1;
	xlSortNormal = 0;	
	xlDown = -4121;   
	xlDatabase = 1; 
	xlLandscape = 2;
	
	WSH = СоздатьОбъект("wscript.shell");   
	
	ПутьКРабочемуСтолу = WSH.ExpandEnvironmentStrings("%USERPROFILE%") + "\Рабочий стол\";
	
	ДатаОтчёта = ТекущаяДата() - 1; 	
	СтрокаЧисло = Строка(ДатаЧисло(ДатаОтчёта));		
	СтрокаМесяц = Строка(ДатаМесяц(ДатаОтчёта));  
	
	Состояние("Формирую отчёт..."); 
	
	// Создаём экземпляр Экселя и обработаем исключение при неудачном запуске---
	Попытка                                              			 //*
		Эксель = СоздатьОбъект("Excel.Application");                     //*
	        //Ограничим книгу одним листом                                   //*
		Эксель.SheetsInNewWorkbook = 1;					 //*
		// Ускорение вывода	                                         //*
		Эксель.DisplayAlerts = False;					 //*
		Эксель.ScreenUpdating = False;                   		 //*
		Эксель.EnableEvents = False;                     		 //*
		Эксель.Visible = False;						 //*
	Исключение 								 //*
		Сообщить(ОписаниеОшибки(), "!!!");               		 //*
		Сообщить("Возможно, MS Excel не установлен.");  		 //*
		Возврат;                                         		 //*
	КонецПопытки;                                        			 //*
	//--------------------------------------------------------------------------    
	
	//-------------------------------------------------------------------------------------
	
	// Добавим новую рабочую книгу
	Книга = Эксель.WorkBooks.Add();  
	
	// Получим окно книги                                     							
	ОкноКниги = Книга.Windows(1);                          	  							
	// Дадим имя окну                                        							
	ОкноКниги.Caption = "Отчёты свеклопункта";             	      
	 
 	//--------------------------------------------------------------------------------	
	
	//-----------------------------------------------------------------------------------  
	// Получим лист
	Лист = Книга.Worksheets(1);
	Лист.Name = "Диспетчеру за " + СтрокаЧисло + "." + СтрокаМесяц; 
	
	Лист.Range("A1").Value = "Район";  
	Лист.Range("B1").Value = "Свеклосдатчик";    
	Лист.Range("C1").Value = "Водитель";      	
	Лист.Range("D1").Value = "АТП"; 	
	Лист.Range("E1").Value = "Гос.№";
	Лист.Range("F1").Value = "Прицеп";
	
	Лист.Range("A1:F1").Font.Size = 9;
	Лист.Range("A1:F1").Font.Bold = True;
	Лист.Range("A1:F1").Font.Name = "Times New Roman"; 
	Лист.Range("A1:F1").HorizontalAlignment = xlHAlignCenter;//выравнивание текста по центру
	
	Для ы=7 по 10 Цикл   	
		Лист.Range("A1").Borders(ы).LineStyle = xlContinuous;
		Лист.Range("B1").Borders(ы).LineStyle = xlContinuous;
		Лист.Range("C1").Borders(ы).LineStyle = xlContinuous;
		Лист.Range("D1").Borders(ы).LineStyle = xlContinuous;
		Лист.Range("E1").Borders(ы).LineStyle = xlContinuous;
		Лист.Range("F1").Borders(ы).LineStyle = xlContinuous;
	КонецЦикла;   
	//--------------------------------------------------------------------------------------------------

	Док = СоздатьОбъект("Документ");
	Док.ИспользоватьЖурнал("ЖурналТТН");
	Док.ВыбратьДокументы(НачалоИнтервала(),КонецИнтервала());   
	
	i=2;
	Пока Док.ПолучитьДокумент() = 1 Цикл
		Если Док.Проведен() = 1 Тогда  
			
			tmp=Строка(Док.Графа("Район")); 
			Лист.Range("A"+i).Value= tmp; 
			
			tmp=Строка(Док.Графа("Свеклосдатчик")); 
			Лист.Range("B"+i).Value= tmp;
			
			tmp=Строка(Док.Графа("АТП")); 
			Лист.Range("D"+i).Value= tmp;
			
			tmp=Строка(Док.Графа("Водитель")); 
			Лист.Range("C"+i).Value= tmp;
			
			tmp=Строка(Док.Графа("НомерАвто")); 
			Лист.Range("E"+i).Value= tmp;
			
			tmp=Строка(Док.Графа("НомерПрицепа")); 
			Лист.Range("F"+i).Value= tmp;
			
			i= i+1;   
		Иначе
			Продолжить;
		КонецЕсли;
	КонецЦикла; 
	//--------------------------------------------------------------------------------------------------------
	
//------------------------------------------------------------------------------------------------	
	//Закрепление области
	Лист.Range("A2").Select();
	ОкноКниги.FreezePanes = True;	
	   
	Лист.Columns("A:C").AdvancedFilter(2, ,Лист.Columns("H:J"), True);
	Лист.Columns("C:D").AdvancedFilter(2, ,Лист.Columns("L:M"), True); 
	
	Лист.Columns("C:C").Cut();
	Лист.Columns("E:E").Insert(xlToRight);
	Эксель.CutCopyMode = False;

	Лист.Range("A1:F1").AutoFilter();
	
	Лист.Range("A1").CurrentRegion.Sort(Лист.Range("D1"), xlAscending, "", ,1, "", 1, xlYes);
	Лист.Range("A1").CurrentRegion.Sort(Лист.Range("C1"), xlAscending, "", ,1, "", 1, xlYes);		
//---------------------------------------------------------------------------------------------------	
	Лист.Range("L1").Select();
	СводнаяПоАТП = Лист.PivotTableWizard;
  
        // Разворачиваем макет сводной таблицы
  	СводнаяПоАТП.SmallGrid = 0;

 	// Теперь разнесем ячейки сводной таблицы
  	СводнаяПоАТП.PivotFields(2).Orientation = 1;  
 	СводнаяПоАТП.PivotFields(1).Orientation = 4;  
		
	
	Лист.Activate();
	Лист.Range("H1").Select();
    	СводнаяПоРайонам = Лист.PivotTableWizard;
    	СводнаяПоРайонам.SmallGrid = 0;

 	СводнаяПоРайонам.PivotFields(1).Orientation = 1;  
 	СводнаяПоРайонам.PivotFields(2).Orientation = 1; 
    	СводнаяПоРайонам.PivotFields(3).Orientation = 4;
	// Где:
	// 1 - Строка
	// 2 - Столбец
	// 3 - Страница
	// 4 - Данные   
	
	//Закрываем панель инструментов сводной таблицы
	Книга.ShowPivotTableFieldList = False;
                     
	Лист.Columns("A:A").ColumnWidth = 16;
	Лист.Columns("B:B").ColumnWidth = 47;
	Лист.Columns("C:C").ColumnWidth = 35;
	Лист.Columns("D:D").ColumnWidth = 22;
	Лист.Columns("E:E").ColumnWidth = 8;
	Лист.Columns("F:F").ColumnWidth = 6;
	
	Лист2 = Книга.Worksheets(2);
	Лист2.Name = "Сводная по районам"; 
	Лист2.Columns("A:A").ColumnWidth = 16;
	
	Лист3 = Книга.Worksheets(1);      	
	Лист3.Name = "Сводная по АТП"; 
	Лист3.Activate(); 

        Лист.Columns("H:M").Delete();
	
	Лист.PageSetup.Orientation = xlLandscape;  
	Лист.PageSetup.LeftMargin = Эксель.InchesToPoints(0.2);
	Лист.PageSetup.RightMargin = Эксель.InchesToPoints(0.2);
	Лист.PageSetup.TopMargin = Эксель.InchesToPoints(0.2);
	Лист.PageSetup.BottomMargin = Эксель.InchesToPoints(0.2);
	
	// Показать границу листа
	Лист.DisplayPageBreaks = True; 
	
    	Книга.SaveAs(ПутьКРабочемуСтолу  + СтрокаМесяц + "." + СтрокаЧисло + "_.xls");

	Эксель.DisplayAlerts = True;
	Эксель.ScreenUpdating = True;
	Эксель.EnableEvents = True;
	Эксель.Visible = True;	

КонецПроцедуры   
