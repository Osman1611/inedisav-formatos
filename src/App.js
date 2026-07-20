import { useState, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";

const DATOS_ESCUELA_INICIAL = {"Ciclo 23 A":{"tutor":"FIDIA FERNANDEZ MELGAREJO","estudiantes":[{"nombre":"ALMANZA PAREJO MICHELLE ZHARICK","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"ARIZA PONCE ZESERGIO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3012651422"},{"nombre":"BARRUETA RESTREPO YESENIA CAROLINA","genero":"F","estado":"MATRICULADO","telefono":"3046257564"},{"nombre":"FERRER LANDERO HIPOLITO JOSE","genero":"M","estado":"MATRICULADO","telefono":"3002403307"},{"nombre":"FLOREZ GARCIA ALEJANDRO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3123668224"},{"nombre":"GUERRERO BOLAÑOS JULIO ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3003622298"},{"nombre":"MENDOZA ALTAMIRANDA YUREM ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"3026698730"},{"nombre":"OLIVAREZ JIMENEZ THALIANA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3023304024"},{"nombre":"PELUFFO MORENO JUAN CARLOS","genero":"M","estado":"MATRICULADO","telefono":"302366968"},{"nombre":"PELUFFO MORENO YERANIA","genero":"F","estado":"MATRICULADO","telefono":"3008649018"},{"nombre":"RAMOS OSORIO DANY DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3003069083"},{"nombre":"RESTREPO BONETT IVANNA MARCELA","genero":"F","estado":"MATRICULADO","telefono":"3262793"},{"nombre":"SABBAGH SICULABA RICARDO DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"SALAZAR VALDES PAMELA DAVISON","genero":"M","estado":"MATRICULADO","telefono":"3042124448"},{"nombre":"TORREGROZA BADEL BRANDON DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3650993"}]},"Ciclo 24 A":{"tutor":"FIDIA FERNANDEZ MELGAREJO","estudiantes":[{"nombre":"ARIZA MONTOYA JOSE MAURICIO","genero":"M","estado":"MATRICULADO","telefono":"3002403307"},{"nombre":"BONILLA LAMPREA BAYRON ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"CABALLERO MEDINA DELKIS FRANCISCO","genero":"M","estado":"MATRICULADO","telefono":"3218436510"},{"nombre":"GARCIA CANTILLO OLIVER JOSUE","genero":"M","estado":"MATRICULADO","telefono":"3014177996"},{"nombre":"GARCIA CUESTA BRAHIAN JOSE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"GUDIÑO RODRIGUEZ STEFANY ESTHER","genero":"F","estado":"MATRICULADO","telefono":"3135056478"},{"nombre":"JIMENEZ HERRERA LILIANA MARIA","genero":"F","estado":"MATRICULADO","telefono":"3043221875"},{"nombre":"MENDEZ ACOSTA LUIS ANGEL","genero":"M","estado":"MATRICULADO","telefono":"3178494070"},{"nombre":"MORENO CONSUEGRA EMANUEL","genero":"M","estado":"MATRICULADO","telefono":"3188165385"},{"nombre":"PIRELA SUAREZ JESUS LIANGEL","genero":"M","estado":"MATRICULADO","telefono":"3043730560"},{"nombre":"RAMOS OSORIO DANIELA MARIA","genero":"F","estado":"MATRICULADO","telefono":"3003069083"},{"nombre":"SAYAS PRENS BRAYAN MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3013164150"},{"nombre":"SIERRA CARRILLO ALANA ISABELA","genero":"F","estado":"MATRICULADO","telefono":"3024519416"},{"nombre":"BOLAÑO LOZANO SOWALDO JESUS0","genero":"M","estado":"MATRICULADO","telefono":"3052810102"},{"nombre":"MENDEZ ACOSTA LUIS ANGEL","genero":"M","estado":"MATRICULADO","telefono":"3026353395"},{"nombre":"SIERRA CARRILLO ALONSO ISABELA","genero":"M","estado":"MATRICULADO","telefono":"3116644489"}]},"Ciclo 25 A":{"tutor":"FIDIA FERNANDEZ MELGAREJO","estudiantes":[{"nombre":"ABREU BARROS ANGELES MAIREL","genero":"F","estado":"MATRICULADO","telefono":"3002637227"},{"nombre":"ALMANZA PAREJO CRIS VALENTINA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"ARIAS BERDUGO LAURA MICHELL","genero":"F","estado":"MATRICULADO","telefono":"3202744603"},{"nombre":"AVENDAÑO CASTILLO YURLEIDYS CAROLINA","genero":"F","estado":"MATRICULADO","telefono":"3013590070"},{"nombre":"BARRERA ARRIETA MICHAEL STIVEN","genero":"M","estado":"MATRICULADO","telefono":"3007773404"},{"nombre":"CAMELO SUAREZ DUBERNEY","genero":"M","estado":"MATRICULADO","telefono":"3024045333"},{"nombre":"CASARES GARCIA SNAIDER DAVID","genero":"M","estado":"MATRICULADO","telefono":"3655791"},{"nombre":"CEDEÑO GONZALEZ JESUS MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3015586935"},{"nombre":"COMENDADOR UTRIA NEIDER JOSE","genero":"M","estado":"MATRICULADO","telefono":"3207067131"},{"nombre":"DUARTE SANZ EMILY DELCARMEN","genero":"F","estado":"MATRICULADO","telefono":"3114047266"},{"nombre":"GARCIA TAPIA GERMAN ANTONIO","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"GOMEZ BORDA CRISTIAN CAMILO","genero":"M","estado":"MATRICULADO","telefono":"3006839774"},{"nombre":"GUDIÑO PADILLA YANDY PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3242951"},{"nombre":"LLINAS AMADOR KEILYN DAYANA","genero":"F","estado":"MATRICULADO","telefono":"3045410895"},{"nombre":"MANJARRES RUIZ MELISSA VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3043595879"},{"nombre":"MEJIA DIAZGRANADOS JESSEIMY MARIA","genero":"F","estado":"MATRICULADO","telefono":"3145648993"},{"nombre":"MIJARES PACHECO MARIA ALEJANDRA","genero":"F","estado":"MATRICULADO","telefono":"3117228775"},{"nombre":"NAVA TERAN CARMEN MARIA","genero":"F","estado":"MATRICULADO","telefono":"3004189929"},{"nombre":"ORTEGA ADARRAGA LUIS MARTIN","genero":"M","estado":"MATRICULADO","telefono":"3006844774"},{"nombre":"OSPINO GONZALEZ JHON DENI","genero":"M","estado":"MATRICULADO","telefono":"3003575032"},{"nombre":"PEREZ DUMAR CAMILO ANDRES","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"PUA GALOFRE JESUS ESTEBAN","genero":"M","estado":"MATRICULADO","telefono":"3113983776"},{"nombre":"RAMIREZ GONZALEZ FABIANA DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3014626733"},{"nombre":"RAMOS CONDE DUVAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3235124278"},{"nombre":"RICO DE LA ROSA BRIAN JOSE","genero":"M","estado":"MATRICULADO","telefono":"3342518"},{"nombre":"MENDOZA ALTAMIRANDA YUREN ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"3026698730"},{"nombre":"CHACON NIÑO LAURA VANESSA","genero":"M","estado":"MATRICULADO","telefono":"3108131190"}]},"Cuarto A":{"tutor":"IVANETH IBAÑEZ ALVARADO","estudiantes":[{"nombre":"ALDANA MEJIA THIAGO JESUS","genero":"M","estado":"MATRICULADO","telefono":"3002769043"},{"nombre":"BARCELO VASQUEZ PAULINA ANAY","genero":"F","estado":"MATRICULADO","telefono":"3007232701"},{"nombre":"BAZA HOYOS ESTEBAN JACOB","genero":"M","estado":"MATRICULADO","telefono":"3015451066"},{"nombre":"CABRERA CAMPO ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":"302861294"},{"nombre":"CASTILLO HERRERA MATTHEWS","genero":"M","estado":"MATRICULADO","telefono":"3003456114"},{"nombre":"CASTRO CASTRO AINHARA ALEXANDRA","genero":"M","estado":"MATRICULADO","telefono":"3023678747"},{"nombre":"CONTRERAS SANCHEZ JAYDEN LIKAN","genero":"M","estado":"MATRICULADO","telefono":"3053224955"},{"nombre":"GARNICA PEDROZA VALERIA ISABEL","genero":"F","estado":"MATRICULADO","telefono":"3504939980"},{"nombre":"GAVIRIA VEGA ABRAHAM","genero":"M","estado":"MATRICULADO","telefono":"3017098437"},{"nombre":"GONZALEZ LLINAS JUANSE RAFAEL","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"GUTIERREZ DEL VALLE STEWIN JOSE","genero":"M","estado":"MATRICULADO","telefono":"3018827498"},{"nombre":"LEON VALDEZ SAMUEL","genero":"M","estado":"MATRICULADO","telefono":"3002512452"},{"nombre":"LOBO ESPAÑA NICOLAS JOSE","genero":"M","estado":"MATRICULADO","telefono":"3015522525"},{"nombre":"LONDOÑO RODAS SANTIAGO","genero":"M","estado":"MATRICULADO","telefono":"3113146989"},{"nombre":"MEDINA NUESI LUCAS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3015684786"},{"nombre":"NAVARRO LARA JESUS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3052847876"},{"nombre":"OSPINO RAMIREZ SAMUEL DE JESUS","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"PERTUZ CAMARGO KEIMIS JOHANA","genero":"F","estado":"MATRICULADO","telefono":"3006414096"},{"nombre":"RIVERA ORTEGA MARIANGEL","genero":"F","estado":"MATRICULADO","telefono":"3003988144"},{"nombre":"RODRIGUEZ ORTEGA LEYNER ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3105472855"},{"nombre":"ROSARIO ORTIZ SEBASTIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3218886047"},{"nombre":"TORTELLO SERRANO ALICIA DIANID","genero":"F","estado":"MATRICULADO","telefono":"3242651918"},{"nombre":"TOVAR PEREZ DAYANIS HELENA","genero":"F","estado":"MATRICULADO","telefono":"3234856176"},{"nombre":"VARGAS ARZUZA ARIADNA MICHELLE","genero":"F","estado":"MATRICULADO","telefono":"3046785492"},{"nombre":"VILLA MENDOZA VANESSA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3003245173"},{"nombre":"VIZCAINO DE LOS REYES LIZ MAIROBYS","genero":"F","estado":"MATRICULADO","telefono":"3244237044"},{"nombre":"ZAGARRA GUTIERREZ YESHUA JAVIER","genero":"M","estado":"MATRICULADO","telefono":"3012837916"},{"nombre":"MARCHENA SANCHEZ ANDRES GUILLERMO","genero":"M","estado":"RETIRADO","telefono":""}]},"Cuarto B":{"tutor":"LUZDARY EGEA CAMARGO","estudiantes":[{"nombre":"ALEAN ITURRIAGO ELIZABETH","genero":"F","estado":"MATRICULADO","telefono":"3007057489"},{"nombre":"ALTAMAR SANCHEZ DANIELA ALEJANDRA","genero":"F","estado":"MATRICULADO","telefono":"3042247178"},{"nombre":"ATENCIO VALBUENA THIAGO EMANUEL","genero":"M","estado":"MATRICULADO","telefono":"3107480919"},{"nombre":"BARRIOS URBINA SAMUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3046663454"},{"nombre":"BRICEÑO MORA ALEJANDRO JOSE","genero":"M","estado":"MATRICULADO","telefono":"3024302452"},{"nombre":"CABALLERO MIRANDA BRYAN SMITH","genero":"M","estado":"MATRICULADO","telefono":"3183684213"},{"nombre":"CACERES MARTELO ISIS ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3204687864"},{"nombre":"COMENDADOR UTRIA JHADIEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3207067131"},{"nombre":"FONTALVO CASSIANI SEBASTIAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3004412922"},{"nombre":"FONTALVO TORRES SARAY","genero":"F","estado":"RETIRADO","telefono":""},{"nombre":"GUTIERREZ TORRES YONANGEL ZAVIER","genero":"M","estado":"MATRICULADO","telefono":"3232280768"},{"nombre":"JARAMILLO CONDE CARLOS ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"3146458139"},{"nombre":"MARIMON PIMENTEL FREN JUNIOR","genero":"M","estado":"MATRICULADO","telefono":"3044103262"},{"nombre":"MARTINEZ ROZO THAEL ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3013738315"},{"nombre":"MERCADO SERRANO GERALDINE","genero":"F","estado":"MATRICULADO","telefono":"3013573426"},{"nombre":"MUÑOZ FIGUEROA YOVANNY JUNIOR","genero":"M","estado":"MATRICULADO","telefono":"3113111026"},{"nombre":"MUÑOZ FRANCO NIDIA LISETH","genero":"F","estado":"MATRICULADO","telefono":"3016973268"},{"nombre":"PEÑARANDA DE LA ROSA MARTIN ELIAS","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"PEREZ FERNANDEZ MATEO","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"RAMIREZ DELGADO DILAN ANDRES","genero":"F","estado":"MATRICULADO","telefono":"3003961958"},{"nombre":"ROLONG OSPINO ARIANA LUCIA","genero":"F","estado":"MATRICULADO","telefono":"3238630454"},{"nombre":"SARMIENTO MIER LITZIS MILAGRO","genero":"F","estado":"MATRICULADO","telefono":"3002351410"},{"nombre":"SOJO HERNANDEZ JEREMY DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3145331516"},{"nombre":"SUAREZ ULLOA JAMES VISQUEL","genero":"M","estado":"MATRICULADO","telefono":"3042093151"},{"nombre":"UTRIA BRICEÑO KIMBERLY MILAGRO","genero":"F","estado":"MATRICULADO","telefono":"3146581471"},{"nombre":"VILLALOBOS RODRIGUEZ YOEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3024393882"},{"nombre":"ZABALETA SANDOVAL SALOME","genero":"F","estado":"MATRICULADO","telefono":"3127515158"},{"nombre":"GONZALEZ SANCHEZ JAIR MATHIAS","genero":"M","estado":"RETIRADO","telefono":"3022760138"}]},"Decimo A":{"tutor":"DEYSI SOFIA RUIZ BARRIOS","estudiantes":[{"nombre":"ANAYA LOPEZ LUIS FERNANDO","genero":"M","estado":"MATRICULADO","telefono":"3013762355"},{"nombre":"ARDILA GARCIA VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3226602157"},{"nombre":"AULAR MEZA DAVID DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3146092839"},{"nombre":"BERMUDEZ SARMIENTO SAMUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"CAMELO SUAREZ DUBEISIS","genero":"F","estado":"MATRICULADO","telefono":"3044058377"},{"nombre":"CASTRO MAURY SANTIAGO DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3245969458"},{"nombre":"CERRA LARA ALLISON DAYANA","genero":"F","estado":"MATRICULADO","telefono":"3017655057"},{"nombre":"COMAS PEREZ FELIPE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"DE LEON ROJAS DANIEL CAMILO","genero":"M","estado":"MATRICULADO","telefono":"3203591173"},{"nombre":"ESLAYT GONZALEZ JACKELINE ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3003325554"},{"nombre":"GONZALEZ CLAVIJO JUAN SEBASTIAN","genero":"M","estado":"MATRICULADO","telefono":"3043194613"},{"nombre":"GUERRA DE LA HOZ JORDAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"HERNANDEZ PEREZ DANNA JAHAZIEL","genero":"M","estado":"MATRICULADO","telefono":"3136186387"},{"nombre":"JAIMES CUEVAS ALDAIR RAFAEL","genero":"M","estado":"MATRICULADO","telefono":"3157982865"},{"nombre":"JARABA RODRIGUEZ NEYMAR","genero":"M","estado":"MATRICULADO","telefono":"3015624465"},{"nombre":"MARTINEZ FONSECA ERIK DAVID","genero":"M","estado":"MATRICULADO","telefono":"3017994213"},{"nombre":"MENDEZ OTERO JEYCAR SOFIA DE LOS ANGELES","genero":"F","estado":"MATRICULADO","telefono":"3023347492"},{"nombre":"OSPINO RAMIREZ MATEO DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3122218402"},{"nombre":"PALACIOS AYALA JONATHAN DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3217182797"},{"nombre":"QUINTERO JIMENEZ JOANYS JOSE","genero":"M","estado":"MATRICULADO","telefono":"3003464476"},{"nombre":"SANCHEZ HERRERA CARLOS ARTURO","genero":"M","estado":"MATRICULADO","telefono":"3173023317"},{"nombre":"VISBAL ESCORCIA DAYDER ANTONIO","genero":"M","estado":"MATRICULADO","telefono":"3146187495"},{"nombre":"ZUÑIGA REBOLLEDO NATALIA","genero":"F","estado":"MATRICULADO","telefono":"3108380684"}]},"Decimo B":{"tutor":"FANNY ISABEL AHUMADA CANTILLO","estudiantes":[{"nombre":"ALVAREZ CHAVARIA MARIANA ESTHER","genero":"F","estado":"MATRICULADO","telefono":"3012869931"},{"nombre":"CAMPO LEONES JUAN SEBASTIAN","genero":"M","estado":"MATRICULADO","telefono":"3012588905"},{"nombre":"DE LAS SALAS ARQUEZ ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"ECHAVEZ GUTIERREZ XAVI MIJAIL","genero":"M","estado":"MATRICULADO","telefono":"3108527286"},{"nombre":"FERRER FARIA YOHANY VERUSKA","genero":"F","estado":"MATRICULADO","telefono":"3043985676"},{"nombre":"GOMEZ OROZCO JAINER RAFAEL","genero":"M","estado":"RETIRADO","telefono":"3005183907"},{"nombre":"GOMEZ OROZCO VIDALBA SOFIA","genero":"F","estado":"RETIRADO","telefono":"3005183907"},{"nombre":"GONZALEZ GONZALEZ MARIA JESUS","genero":"F","estado":"MATRICULADO","telefono":"3013810427"},{"nombre":"HINCAPIE MARQUEZ MATHIAS DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3462362"},{"nombre":"JIMENEZ POLO DUVAN ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"6053709887"},{"nombre":"MANJARRES SUAREZ ELIANNYS CAROLINA","genero":"F","estado":"MATRICULADO","telefono":"3012038925"},{"nombre":"MARTINEZ BERDUGO JUNIOR ALONSO","genero":"M","estado":"MATRICULADO","telefono":"3022926268"},{"nombre":"MARTINEZ FONSECA SEBASTIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3017994213"},{"nombre":"MAYORGA CERVANTES JULIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3008071766"},{"nombre":"OCHOA GRACIANI SAMUEL DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3164551772"},{"nombre":"OROZCO ORTEGA DYLAN JOSE","genero":"M","estado":"MATRICULADO","telefono":"3014145188"},{"nombre":"OTERO VELASQUEZ DAMIAN JESUS","genero":"M","estado":"MATRICULADO","telefono":"3010669059"},{"nombre":"OYOLA DURAN KEIDY VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3206454"},{"nombre":"PADILLA TORNE KEVIN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3103564998"},{"nombre":"RAMIREZ PADILLA JULIANA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"REYES CHIRINOS JOSLIANNY GUADALUPE","genero":"F","estado":"MATRICULADO","telefono":"3015332819"},{"nombre":"RODRIGUEZ MIRANDA MARIA FERNANDA","genero":"F","estado":"MATRICULADO","telefono":"3126133479"},{"nombre":"ROMERO ORTIZ SANTIAGO MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3128191484"},{"nombre":"SANCHEZ HERNANDEZ SHAROM SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3203743107"},{"nombre":"SANDOVAL LUGO STEVEN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3043959906"},{"nombre":"SARMIENTO LACERA SEBASTIAN","genero":"M","estado":"MATRICULADO","telefono":"3146299502"}]},"Jardin A":{"tutor":"NELDA FONTALVO BARANDICA","estudiantes":[{"nombre":"BARRETO CARO VALERIA","genero":"F","estado":"MATRICULADO","telefono":"3183689299"},{"nombre":"BARRIOS BERMUDES ARATH JAEL","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"BARRIOS OROZCO ZARA VICTORIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"BUELVAS CABARCAS SHADYA LUCIA","genero":"F","estado":"MATRICULADO","telefono":"3013358412"},{"nombre":"CAMARGO BOZO LIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"ESCORCIA BEITAR HANNAH SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3148592597"},{"nombre":"FAJARDO CARDOZO LIAM DAVID","genero":"M","estado":"MATRICULADO","telefono":"3012160174"},{"nombre":"GONZALEZ DIAZ SANTIAGO JADDITH","genero":"M","estado":"MATRICULADO","telefono":"3152068651"},{"nombre":"GUERRERO JARAMILLO MATEO DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3172497752"},{"nombre":"HERNANDEZ CONTRERAS TEO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3024762978"},{"nombre":"IBARRA CASTILLO STEVEN DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"NAVARRO PEREZ LUIS ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":"3016188493"},{"nombre":"ORTEGA ACOSTA JADIEL XAVI","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"PALOMO ARAUJO LAUREN SOFIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"SALAS GUERRA DORIS SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3239607813"},{"nombre":"SANCHEZ JIMENEZ NOAH STEVEN","genero":"M","estado":"MATRICULADO","telefono":"3006125862"},{"nombre":"TORRADO SUAREZ ASHLY MILAGROS","genero":"F","estado":"MATRICULADO","telefono":"3025433440"},{"nombre":"TORREGROZA ALMANZA IAM JOSUE","genero":"M","estado":"MATRICULADO","telefono":"3004446200"},{"nombre":"TORRES RADA SANTIAGO JESUS","genero":"M","estado":"MATRICULADO","telefono":"3016899359"},{"nombre":"ABELLO GUTIERREZ DALIA VICTORIA","genero":"F","estado":"RETIRADO","telefono":"3015207171"}]},"Noveno A":{"tutor":"JUAN ANTONIO CONSTANTE CASTAÑEDA","estudiantes":[{"nombre":"ARIAS PIMENTEL BREYNER DAVID","genero":"M","estado":"MATRICULADO","telefono":"3044103262"},{"nombre":"BRAYS CORRALES DANNA MICHELLE","genero":"F","estado":"MATRICULADO","telefono":"3053036161"},{"nombre":"CABALLERO MC LEAN LUDWIG","genero":"M","estado":"MATRICULADO","telefono":"3187388067"},{"nombre":"FERRER FARIA YEREMY ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":"3043985676"},{"nombre":"FLORES RODRIGUEZ RONNIANLYS ESTEFANIA","genero":"F","estado":"MATRICULADO","telefono":"3045748508"},{"nombre":"HERNANDEZ OCHOA JUAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3177326703"},{"nombre":"HERNANDEZ OLIVEROS LAURA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3001644899"},{"nombre":"IBARRA NAVA ASHLEY NAZARET","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"LOPEZ CABALLERO ADONYS JAVIER","genero":"M","estado":"MATRICULADO","telefono":"3216427561"},{"nombre":"LOPEZ CASTELLAR MARIANGEL","genero":"F","estado":"MATRICULADO","telefono":"3005121739"},{"nombre":"MARTINEZ GOMEZ ANDRES FELIPE","genero":"M","estado":"MATRICULADO","telefono":"3106461933"},{"nombre":"MARTINEZ GOMEZ VALENTINA PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3014595967"},{"nombre":"MEZA MARQUEZ WILMAR SANTIAGO","genero":"M","estado":"MATRICULADO","telefono":"3183441065"},{"nombre":"MIRANDA MENDOZA MATTHEW JOSE","genero":"M","estado":"MATRICULADO","telefono":"3014272169"},{"nombre":"NARVAEZ VELASQUEZ ISAAC DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3177557567"},{"nombre":"PADILLA HERNANDEZ YIRAIDYS DE LOS ANGELES","genero":"F","estado":"MATRICULADO","telefono":"3012097469"},{"nombre":"RODRIGUEZ HERNANDEZ DENNINSON JESUS","genero":"M","estado":"MATRICULADO","telefono":"3017981695"},{"nombre":"ROJANO AGUILAR SEBASTIAN FABIAN","genero":"M","estado":"MATRICULADO","telefono":"3166210577"},{"nombre":"SANCHEZ ESCORCIA SHARON MARIA","genero":"F","estado":"MATRICULADO","telefono":"3246012585"},{"nombre":"SOLANO CABALLERO DIEGO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3105675465"},{"nombre":"VERGARA FERNANDEZ GABRIEL DAVID","genero":"F","estado":"MATRICULADO","telefono":"3006301952"},{"nombre":"VISBAL LEON KIMBERLY JULIETH","genero":"F","estado":"MATRICULADO","telefono":"3214386311"},{"nombre":"YEPES RESTREPO LUIS EDUARDO","genero":"M","estado":"MATRICULADO","telefono":"3007240257"},{"nombre":"POLO OLMOS JUAN MIGUEL","genero":"M","estado":"RETIRADO","telefono":"3046372843"}]},"Noveno B":{"tutor":"LOTTY JUDITH CAMARGO PEÑALOZA","estudiantes":[{"nombre":"ABRIL DE LA HOZ JHON SEBASTIAN","genero":"M","estado":"MATRICULADO","telefono":"3117579689"},{"nombre":"AHUMADA ARIZA YAKELI MICHEL","genero":"F","estado":"MATRICULADO","telefono":"3003024630"},{"nombre":"BUSTAMANTE PALENCIA ANDRES DAVID","genero":"M","estado":"MATRICULADO","telefono":"3145456268"},{"nombre":"CABALLERO MC LEAN LUSHEILA","genero":"F","estado":"MATRICULADO","telefono":"3187388067"},{"nombre":"CARO MESTRA DANNA CECILIA","genero":"F","estado":"MATRICULADO","telefono":"3004255829"},{"nombre":"CASTELLAR NAVAS THALIANY DISHEIDY","genero":"F","estado":"MATRICULADO","telefono":"3014805634"},{"nombre":"CORONADO AGRESOTH VANESA PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3022707049"},{"nombre":"DAZA VEGA OLGA MICHELLE","genero":"F","estado":"MATRICULADO","telefono":"3013213250"},{"nombre":"DE LA CRUZ DE LOS REYES SAMUEL MAURICIO","genero":"M","estado":"MATRICULADO","telefono":"3138409939"},{"nombre":"DE LA HOZ LLERENA JOSEPH ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"3014519459"},{"nombre":"FONTALVO TORRES JOHAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3246731589"},{"nombre":"GONZALEZ MOLINA ASHLEY NICOLL","genero":"F","estado":"MATRICULADO","telefono":"3007746868"},{"nombre":"GUZMAN HERNANDEZ MARIA MIEL","genero":"F","estado":"MATRICULADO","telefono":"3226401994"},{"nombre":"INSIGNARES MARTINEZ DARWIN STIVEN","genero":"M","estado":"MATRICULADO","telefono":"3005106597"},{"nombre":"JURADO CARDENAS LIZZY FABIANA","genero":"F","estado":"MATRICULADO","telefono":"3004935124"},{"nombre":"MARTINEZ FONSECA JUAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3017994213"},{"nombre":"MEDRANO CHARRYS DANIEL JOSUE","genero":"M","estado":"MATRICULADO","telefono":"3124870"},{"nombre":"MEZA RODRIGUEZ MATHIAS SANTIAGO","genero":"M","estado":"MATRICULADO","telefono":"3006124325"},{"nombre":"MOTTA MONTAÑO YOSUE GABRIEL","genero":"M","estado":"MATRICULADO","telefono":"3015204999"},{"nombre":"RIPOLL MEDRANO SHEILYS MICHELL","genero":"F","estado":"MATRICULADO","telefono":"3046515240"},{"nombre":"RODRIGUEZ DE ALBA ANGEL ERNESTO","genero":"M","estado":"MATRICULADO","telefono":"3126701554"},{"nombre":"RODRIGUEZ DE LOS REYES ROBERTH STEVEN","genero":"M","estado":"MATRICULADO","telefono":"3136959113"},{"nombre":"VESGA LOZANO JUAN FELIPE","genero":"M","estado":"MATRICULADO","telefono":"3015383079"},{"nombre":"YEPES RESTREPO LUIS SEBASTIAN","genero":"M","estado":"MATRICULADO","telefono":"3007240257"}]},"Octavo A":{"tutor":"NOELVIS JOSEFINA MENDEZ ALVAREZ","estudiantes":[{"nombre":"ALCAZAR JIMENEZ SANTIAGO JOSE","genero":"M","estado":"MATRICULADO","telefono":"3005200561"},{"nombre":"BARCELO VASQUEZ SUSANA EDITH","genero":"F","estado":"MATRICULADO","telefono":"3007232701"},{"nombre":"BAYONA CARMONA SARA VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3107059968"},{"nombre":"BECERRA PEREZ MIGUEL JOSE","genero":"M","estado":"MATRICULADO","telefono":"3016703305"},{"nombre":"CABARCAS PAEZ JEREMY JUNIOR","genero":"M","estado":"MATRICULADO","telefono":"3007244587"},{"nombre":"CHALELA BELLO ISAIAS","genero":"M","estado":"MATRICULADO","telefono":"3008732348"},{"nombre":"CHAMORRO CONSTANTE DANILETH MARIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"ECHEVERRIA MCAUSLAND GIOVANNI ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3017770553"},{"nombre":"FLORES MALDONADO SEBASTIAN AUGUSTO","genero":"M","estado":"MATRICULADO","telefono":"3116710445"},{"nombre":"GOMEZ ACOSTA SARAI ESTER","genero":"F","estado":"MATRICULADO","telefono":"3012851188"},{"nombre":"GONZALEZ VILLARROEL DIUVERLYS VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3127120530"},{"nombre":"IMITOLA VALDIVIEZO GILBER RAFAEL","genero":"M","estado":"MATRICULADO","telefono":"3023791672"},{"nombre":"JARAMILLO CONDE ELISA MARCELA","genero":"F","estado":"MATRICULADO","telefono":"3235124278"},{"nombre":"JULIO TORRENEGRA KAROLL DAYANA","genero":"F","estado":"MATRICULADO","telefono":"3638131"},{"nombre":"LARA PADILLA DEREK JETER","genero":"M","estado":"MATRICULADO","telefono":"3203936612"},{"nombre":"LOPEZ LOPEZ DILAN JESUS","genero":"M","estado":"MATRICULADO","telefono":"3234702779"},{"nombre":"OSORIO QUINTANA GERALDIN","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"PALLARES PEREIRA CAMILO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3127388315"},{"nombre":"PARRA RODRIGUEZ MARIANA","genero":"F","estado":"MATRICULADO","telefono":"3244054278"},{"nombre":"PEREZ ESTRADA YORDI SAMUEL","genero":"M","estado":"MATRICULADO","telefono":"3215071584"},{"nombre":"QUINTERO ORTIZ MIGUEL ANGEL","genero":"M","estado":"MATRICULADO","telefono":"3116027128"},{"nombre":"RAMOS ALVAREZ ANDRES FELIPE","genero":"M","estado":"MATRICULADO","telefono":"3043794975"},{"nombre":"RAMOS GONZALEZ JHONATAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3014238538"},{"nombre":"RODRIGUEZ RAMOS EDUARDO MIGUEL","genero":"M","estado":"MATRICULADO","telefono":"3113570810"},{"nombre":"ROMERO PEÑA SALOMON ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"SANTAMARIA BOTELLO ALEXANDER ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"3207315122"},{"nombre":"SANTIAGO DE LAS SALAS STIWAR DAVID","genero":"M","estado":"MATRICULADO","telefono":"3014624969"},{"nombre":"SARMIENTO CABADIA YEIMER ELIAS","genero":"M","estado":"MATRICULADO","telefono":"3023954559"},{"nombre":"SERRANO VERA JOSUE ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":"3044777536"},{"nombre":"SOJO HERNANDEZ CAMILA ARIANI","genero":"F","estado":"MATRICULADO","telefono":"3145331516"},{"nombre":"SOTO VILORIA JOSUE DAVID","genero":"M","estado":"MATRICULADO","telefono":"3043175331"},{"nombre":"WALTEROS GONZALEZ JENIFER ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3013810427"}]},"Octavo B":{"tutor":"ZULIMA LUCIA RODRIGUEZ CARRILLO","estudiantes":[{"nombre":"ALTAMAR JULIO MATIAS JOSE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"AVENDAÑO OSORIO EMILLY SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3046357498"},{"nombre":"AVENDAÑO OSORIO IVANNA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3046357498"},{"nombre":"BALLESTA LORA JOAM","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"BARRIOS AVENDAÑO ALDRY JOSUE","genero":"M","estado":"MATRICULADO","telefono":"3014867877"},{"nombre":"BARRIOS MOSCOTE JUAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3043545291"},{"nombre":"BELLO ORDOÑEZ MARIA CAMILA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"CARREÑO ARZUZA JORGE DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"DEL TORO CASTRO LUIS MIGUEL","genero":"M","estado":"MATRICULADO","telefono":"3017578927"},{"nombre":"DIAZ VILLASMIL ANDERSON ADRIAN","genero":"M","estado":"MATRICULADO","telefono":"3188371880"},{"nombre":"DITA CORREA LITZY ALEXANDRA","genero":"F","estado":"MATRICULADO","telefono":"3015320424"},{"nombre":"FANDIÑO MORALES BRIANY SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3042995293"},{"nombre":"GARCIA MOLINA SAMANTHA","genero":"F","estado":"MATRICULADO","telefono":"3006550185"},{"nombre":"GONZALEZ AVENDAÑO CHRISTOPHER","genero":"M","estado":"MATRICULADO","telefono":"3014867877"},{"nombre":"JIMENEZ CAMARGO DANIELA PATRICIA","genero":"F","estado":"MATRICULADO","telefono":"3014632976"},{"nombre":"LORA ESCALANTE CHRISTIAN CAMILO","genero":"M","estado":"MATRICULADO","telefono":"3103598194"},{"nombre":"MONTERROZA AHUMADA MARIA JOSE","genero":"F","estado":"MATRICULADO","telefono":"3003024630"},{"nombre":"MUÑOZ MALDONADO MAIKEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3154901148"},{"nombre":"PACHECO ELLES HENRY DAVID","genero":"M","estado":"MATRICULADO","telefono":"3013535300"},{"nombre":"PRADA GARCIA DONOVAN","genero":"M","estado":"MATRICULADO","telefono":"3023025973"},{"nombre":"RUSS VEGA MICHELL ALEXANDRA","genero":"F","estado":"MATRICULADO","telefono":"3013360397"},{"nombre":"SANCHEZ SIERRA KENITH DAVID","genero":"M","estado":"MATRICULADO","telefono":"3013224579"},{"nombre":"SANDOVAL FLOREZ ALLAMS DARIO","genero":"M","estado":"MATRICULADO","telefono":"3206139243"},{"nombre":"SANDOVAL LUGO YENILETH VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3015981055"},{"nombre":"SANTAMARIA ORTIZ DILAN","genero":"M","estado":"MATRICULADO","telefono":"3122850193"},{"nombre":"TAPIAS CANO KEVIN DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3004530118"},{"nombre":"TERAN TRUJILLO YEISUS ALBERTO","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"TOVAR CHICA MARIA VICTORIA","genero":"F","estado":"MATRICULADO","telefono":"3187642309"},{"nombre":"VALENZUELA ROMERO RACHELL MARIA","genero":"F","estado":"MATRICULADO","telefono":"3008476152"},{"nombre":"VILLALOBOS RODRIGUEZ CAMILA VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3015587770"},{"nombre":"SIERRA MACHADO ELIANA SOFIA","genero":"F","estado":"RETIRADO","telefono":"3045515353"}]},"Primero A":{"tutor":"LUZ ELENA DE LA HOZ CUENTAS","estudiantes":[{"nombre":"ARAUJO JULIO JUAN DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3107243664"},{"nombre":"ARTEAGA LEON LIOMAR ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3026153379"},{"nombre":"BARRAZA VASQUEZ DANLYS JESUS","genero":"M","estado":"MATRICULADO","telefono":"3005583434"},{"nombre":"BARRIOS OROZCO ANGEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"321756454"},{"nombre":"BOVEA VASQUEZ MARYCHELL SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3011678864"},{"nombre":"CAFIEL DE MOYA SALIM JAMID","genero":"M","estado":"MATRICULADO","telefono":"3026518658"},{"nombre":"CASTELLAR NAVAS ISABELLA","genero":"F","estado":"MATRICULADO","telefono":"3245810316"},{"nombre":"CASTILLEJO ROCHA JESUS ELIAS","genero":"M","estado":"MATRICULADO","telefono":"3046664079"},{"nombre":"CASTILLO MEZA ASHLAM ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3013246311"},{"nombre":"MEJIA ARIAS ALAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3052835961"},{"nombre":"MENDOZA DIAZ MILAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3238085123"},{"nombre":"MUÑOZ COLINA EYTHAN SMITH","genero":"M","estado":"RETIRADO","telefono":"3186410802"},{"nombre":"MUÑOZ MALDONADO MATEO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3013073764"},{"nombre":"PEREZ GUERRERO JUAN SEBASTIAN","genero":"M","estado":"MATRICULADO","telefono":"3046498468"},{"nombre":"RAMOS OSORIO ABRAHAM DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3005756321"},{"nombre":"RAMOS OSORIO SALOME DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3005756312"},{"nombre":"RIPOLL PEREZ JACOB ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3015695245"},{"nombre":"SILVERA MENDEZ ISABELLA","genero":"F","estado":"MATRICULADO","telefono":"3169349733"},{"nombre":"TORRES SARMIENTO SARAH SOPHIA","genero":"F","estado":"MATRICULADO","telefono":"3012975283"},{"nombre":"VEGA MONTALVO LEANDRO EMANUEL","genero":"M","estado":"MATRICULADO","telefono":"3004831774"},{"nombre":"BALLESTA LORA JANDEL","genero":"M","estado":"RETIRADO","telefono":"3012698415"}]},"Primero B":{"tutor":"NORA GONZALEZ RODRIGUEZ","estudiantes":[{"nombre":"ANDRADE BARRANCO KYLIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3042199728"},{"nombre":"BUENO LASCANO JOEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3001736383"},{"nombre":"CONRADO PEREZ DILAN JESUS","genero":"M","estado":"MATRICULADO","telefono":"3243577450"},{"nombre":"DOMINGUEZ DE ALBA SALOME","genero":"F","estado":"MATRICULADO","telefono":"3004091343"},{"nombre":"ESCORCIA CAÑIZALEZ NEITHAN DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"GIL YUDEX WILLIAM JUNIOR","genero":"M","estado":"MATRICULADO","telefono":"3207731452"},{"nombre":"HERNANDEZ ROLONG JESUS VALENTIN","genero":"M","estado":"MATRICULADO","telefono":"3007170515"},{"nombre":"HOYOS SARMIENTO LIONEL ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3004218885"},{"nombre":"JIMENEZ PACHECO ABIGAIL","genero":"F","estado":"MATRICULADO","telefono":"3017995054"},{"nombre":"LEON BORGES DARNELYS VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3225249118"},{"nombre":"MUÑOZ COLINA KATTALEYA IVETT","genero":"F","estado":"RETIRADO","telefono":"3186410802"},{"nombre":"NAVA TERAN NAIROBYS DAYERLYS","genero":"F","estado":"MATRICULADO","telefono":"3004189929"},{"nombre":"PAREJO GUANAREZ EMANUEL JOSE","genero":"M","estado":"MATRICULADO","telefono":"3122287802"},{"nombre":"RODRIGUEZ DIAZ EMMANUEL","genero":"M","estado":"MATRICULADO","telefono":"3043609392"},{"nombre":"SAAVEDRA GARNICA ABRAHAM DAVID","genero":"M","estado":"MATRICULADO","telefono":"3504939980"},{"nombre":"SANJUAN ANGULO MATIAS","genero":"M","estado":"MATRICULADO","telefono":"3006560694"},{"nombre":"VILLA TORREGROZA MATHIAS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3016158875"},{"nombre":"WALTEROS GONZALEZ IVETH MARIA","genero":"F","estado":"MATRICULADO","telefono":"3013810427"},{"nombre":"GONZALES SANCHEZ JAYLIN","genero":"F","estado":"RETIRADO","telefono":""}]},"Quinto A":{"tutor":"ROCIO MUÑOZ SABALZA","estudiantes":[{"nombre":"AGUILERA FERNADEZ SANTIAGO ISAAC","genero":"M","estado":"MATRICULADO","telefono":"3045990457"},{"nombre":"BARRAZA GARIZABALO EDGARDO MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3137159798"},{"nombre":"BUELVAS CABARCAS TANIA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3013358412"},{"nombre":"CASTRO ALVAREZ NEYMAR DUBAN","genero":"M","estado":"MATRICULADO","telefono":"3016136823"},{"nombre":"ESCOBAR CELIN LUIS SNEYDER","genero":"M","estado":"MATRICULADO","telefono":"3152505457"},{"nombre":"FONTALVO OROZCO MERY PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3043286473"},{"nombre":"FONTALVO TORRES ISAAC DAVID","genero":"M","estado":"RETIRADO","telefono":""},{"nombre":"GARNICA PEDROZA DYLAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3504939980"},{"nombre":"GOMEZ PINTO FABIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3002465332"},{"nombre":"GONZALEZ MELENDEZ AURA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3016261417"},{"nombre":"HERNANDEZ PADILLA JOSE ANGEL","genero":"M","estado":"MATRICULADO","telefono":"3218844757"},{"nombre":"HERRERA OBESO GABRIELA ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3008131870"},{"nombre":"IBARRA CASTILLO IVANNA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3013052519"},{"nombre":"JIMENEZ CASTILLA GABRIELA LUCIA","genero":"F","estado":"MATRICULADO","telefono":"3016709658"},{"nombre":"LOBO MARTINEZ FABIAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3012722366"},{"nombre":"MANJARRES ILLERA IAN SAID","genero":"M","estado":"MATRICULADO","telefono":"3012470205"},{"nombre":"MARIMON PIMENTEL JULEIDYS SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3244494696"},{"nombre":"MARTINEZ TARAZONA SAMUEL ELIAS","genero":"M","estado":"MATRICULADO","telefono":"3127076181"},{"nombre":"MENDOZA DIAZ LUCIANA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3238085123"},{"nombre":"MORALES JULIO KARIANNYS SOFIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"NEMEZ FIGUEROA BRIGITTE SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3013907791"},{"nombre":"OLIVERO ALVAREZ SEBASTIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3022048063"},{"nombre":"OSAL TORRES CAMILA SOPHIA","genero":"F","estado":"MATRICULADO","telefono":"3015790038"},{"nombre":"RODRIGUEZ DE ALBA CARLOS SANTIAGO","genero":"M","estado":"MATRICULADO","telefono":"3004091346"},{"nombre":"RODRIGUEZ GARCIA JAVID ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"3017673750"},{"nombre":"ROMERO CRISTANCHO FRANCISCO ELIAN","genero":"M","estado":"MATRICULADO","telefono":"3177588245"},{"nombre":"RUIZ HOUSSET KEINER","genero":"M","estado":"MATRICULADO","telefono":"3024251542"},{"nombre":"SAMPAYO MARTINEZ LIAM ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3007010427"},{"nombre":"SARMIENTO CABADIA THIAGO ESTEBAN","genero":"M","estado":"MATRICULADO","telefono":"3023954559"},{"nombre":"SIERRA CARRILLO JHOSUA JOSUE","genero":"M","estado":"MATRICULADO","telefono":"3135795507"},{"nombre":"UTRIA ALVAREZ SANTIAGO JOSE","genero":"M","estado":"MATRICULADO","telefono":"3012205758"}]},"Quinto B":{"tutor":"BERENICE MERCADO BORJA","estudiantes":[{"nombre":"AGUILAR ACEVEDO SANTIAGO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3005308797"},{"nombre":"ALVAREZ CHAVARRIA SAMUEL ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"ARRIETA CANTILLO MARLON JUNIOR","genero":"M","estado":"MATRICULADO","telefono":"3044187381"},{"nombre":"BARRIOS CARDENAS EMANUEL ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3215776073"},{"nombre":"BARRIOS TALERO SAMUEL JOSE","genero":"M","estado":"MATRICULADO","telefono":"3205501141"},{"nombre":"CAMARGO GONZALEZ JOSE JUNIOR","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"CANTILLO ESCALANTE JEFFREY SMITH","genero":"M","estado":"MATRICULADO","telefono":"3043403644"},{"nombre":"CONTRERAS SANCHEZ WENDY NICOLL","genero":"F","estado":"MATRICULADO","telefono":"3227980101"},{"nombre":"CORONELL GAMARRA DAPHNE MICHELLE","genero":"F","estado":"MATRICULADO","telefono":"3016135397"},{"nombre":"COTES LOPEZ AUSTIN DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3126363501"},{"nombre":"ESCORCIA GARCIA VALERIA","genero":"F","estado":"MATRICULADO","telefono":"3002491149"},{"nombre":"FONTALVO TORRES SAMUEL DAVID","genero":"M","estado":"RETIRADO","telefono":""},{"nombre":"GONZALEZ AVENDAÑO JOSHUA","genero":"M","estado":"MATRICULADO","telefono":"3014867877"},{"nombre":"IZQUIERDO BOTELLO JESUS GABRIEL","genero":"M","estado":"MATRICULADO","telefono":"3015424303"},{"nombre":"JARAMILLO CONDE CARLOS MARIO","genero":"M","estado":"MATRICULADO","telefono":"3146458139"},{"nombre":"MARTINEZ SUAREZ NELIANNYS PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3042093151"},{"nombre":"MENDOZA PEREZ ARANZA SOFYA","genero":"F","estado":"MATRICULADO","telefono":"3128767277"},{"nombre":"PEDEAÑA NARANJO THAELL MARIAN","genero":"F","estado":"MATRICULADO","telefono":"3142939511"},{"nombre":"RAMIREZ CARRANZA SEBASTIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3184491743"},{"nombre":"RAMIREZ PEDROZA MARTIN DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3015157422"},{"nombre":"RIVERA MONTERO BREIDIS ELIZA","genero":"F","estado":"MATRICULADO","telefono":"3045566675"},{"nombre":"SALAZAR GOMEZ DANIEL ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":"3172593016"},{"nombre":"SANABRIA CARRACEDO ISABELLA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"SANCHEZ BLANCO THIAGO ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3043619144"},{"nombre":"SANTAMARIA ORTIZ EMILY","genero":"M","estado":"MATRICULADO","telefono":"3122580266"},{"nombre":"SERPA LOPEZ ANILYS ESTHER","genero":"F","estado":"MATRICULADO","telefono":"3235484617"},{"nombre":"STELL PATIÑO THIAGO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3017749868"},{"nombre":"SUAREZ GELIZ SALOME","genero":"F","estado":"MATRICULADO","telefono":"3212840918"},{"nombre":"TORRES LOPEZ ZOE SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3246009902"},{"nombre":"VILLALOBOS BELTRAN SHARITH SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3013993615"}]},"Segundo A":{"tutor":"CARMEN AHUMADA OLIVARES","estudiantes":[{"nombre":"ALVARINO ALTAMAR ASHLEY SOFIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"AMAYA PACHON MATIAS","genero":"M","estado":"MATRICULADO","telefono":"3242374939"},{"nombre":"ARRIETA CANTILLO FERNANDO JOSE","genero":"M","estado":"MATRICULADO","telefono":"3004223715"},{"nombre":"CAMACHO PUA ALEXANDER DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3033797869"},{"nombre":"CAMARGO GONZALEZ GELENIS THALIA","genero":"M","estado":"MATRICULADO","telefono":"3244493709"},{"nombre":"CANTILLO ESCALANTE MAILETH SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3004235768"},{"nombre":"CARABALLO NUÑEZ YUSETH DAVID","genero":"M","estado":"MATRICULADO","telefono":"3107343373"},{"nombre":"CISNERO RINCON NAILETH VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3012652329"},{"nombre":"CORONELL GAMARRA DASHIELL PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3016135397"},{"nombre":"CORRALES JIMENEZ ARIADNA LUCIA","genero":"F","estado":"MATRICULADO","telefono":"3022936702"},{"nombre":"CORREA HOYOS LUCIANA MARGARITA","genero":"F","estado":"MATRICULADO","telefono":"3017765384"},{"nombre":"COTES LOPEZ CARLOS ANDRES","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"DEL VALLE GUZMAN DAYMER ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3018827498"},{"nombre":"FERRER FARIA YARETH DAVID","genero":"M","estado":"MATRICULADO","telefono":"3043985676"},{"nombre":"GIL YUDEX MILAGROS DEL CARMEN","genero":"F","estado":"MATRICULADO","telefono":"3006949546"},{"nombre":"GOMEZ MARTINEZ NICHOLAS ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3004041789"},{"nombre":"GONZALEZ VARGAS DIUVER ALFREDO","genero":"M","estado":"MATRICULADO","telefono":"3127120530"},{"nombre":"GORDON CASTRO EIDAN ALI","genero":"M","estado":"MATRICULADO","telefono":"3188826323"},{"nombre":"IBAÑEZ VALLEJO EMMANUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3226360233"},{"nombre":"INSIGNARES REYES MAYTE ANGELIN","genero":"F","estado":"MATRICULADO","telefono":"3042722887"},{"nombre":"LLANOS MONTENEGRO MOISES SALOMON","genero":"M","estado":"MATRICULADO","telefono":"3147974344"},{"nombre":"LOPEZ NUÑEZ SAMUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3215644426"},{"nombre":"MARIN CAÑATE KEILIZ ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3242504384"},{"nombre":"MERCADO MEZA SHADAY SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3116392084"},{"nombre":"MERCADO VILLEGAS OMAR ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3233941696"},{"nombre":"MOLINA ROMERO BRAYER","genero":"M","estado":"MATRICULADO","telefono":"3023565902"},{"nombre":"PACHECO YANES JESUS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3002751002"},{"nombre":"RAMOS SUAREZ ABRAHAM DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3012038925"},{"nombre":"RIVERA ORTEGA MIGUEL ANGEL","genero":"M","estado":"MATRICULADO","telefono":"3024479759"},{"nombre":"RODRIGUEZ GARCIA MIGUEL ANTONIO","genero":"M","estado":"MATRICULADO","telefono":"3043193265"},{"nombre":"RODRIGUEZ MIRANDA ANDREY CAMILO","genero":"M","estado":"MATRICULADO","telefono":"3117071162"},{"nombre":"ROJANO LIDUEÑAS CHRISTOPHER","genero":"M","estado":"MATRICULADO","telefono":"3009441109"},{"nombre":"RONDON VEGA GENESIS SOFIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"SANCHEZ CONTRERAS FABIAN MISAEL","genero":"M","estado":"MATRICULADO","telefono":"3007168580"},{"nombre":"SANCHEZ ESCOBAR CAMILO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3003423553"},{"nombre":"SANTIAGO PEÑA MANUEL DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3016151149"},{"nombre":"TOVAR PEREZ DAVID SANTIAGO","genero":"M","estado":"MATRICULADO","telefono":"3205507885"},{"nombre":"DONADO SARMIENTO GERARD JOSE","genero":"M","estado":"RETIRADO","telefono":"3015947944"}]},"Septimo A":{"tutor":"CLARIBEL CELESTE REALES OJEDA","estudiantes":[{"nombre":"ARRIETA CASTILLO YEIMAR JESUS","genero":"M","estado":"MATRICULADO","telefono":"3149743476"},{"nombre":"BARRANCO FRANCO ARIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3016254171"},{"nombre":"BARRIOS RODRIGUEZ ALEX DAVID","genero":"M","estado":"MATRICULADO","telefono":"3044630884"},{"nombre":"CAMPO OROZCO NEYMAR","genero":"M","estado":"MATRICULADO","telefono":"3116716834"},{"nombre":"CASTRO MOSCOTE SEBASTIAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3234638300"},{"nombre":"CAUSADO VALDES ARIANA FERNANDA","genero":"F","estado":"MATRICULADO","telefono":"3236333134"},{"nombre":"CELIN CHIRINOS WILFER JOSE","genero":"M","estado":"MATRICULADO","telefono":"3022310484"},{"nombre":"DE LA HOZ GONZALEZ ANDRES DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"FLOREZ SANDOVAL YERAY MATIAS","genero":"M","estado":"MATRICULADO","telefono":"3114260209"},{"nombre":"FONTALVO PAEZ WULIANGELA ESTEFANI","genero":"F","estado":"MATRICULADO","telefono":"3136275438"},{"nombre":"GARCIA ARAGON HECTOR EDUARDO","genero":"M","estado":"MATRICULADO","telefono":"3007448623"},{"nombre":"GONZALEZ BERROCAL BLADIMIR ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3022802591"},{"nombre":"GUERRA VILLARREAL KAEL EDUARDO","genero":"M","estado":"MATRICULADO","telefono":"3017382136"},{"nombre":"GUETTE ACUÑA JAIME DAVID","genero":"M","estado":"MATRICULADO","telefono":"3017641574"},{"nombre":"HERNANDEZ NAVAJA JESUS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3045309893"},{"nombre":"GUTIERREZ CRECIENTE MARIANGEL","genero":"F","estado":"MATRICULADO","telefono":"3001502672"},{"nombre":"JIMENEZ MARTINEZ YULEIMY CAMILA","genero":"F","estado":"MATRICULADO","telefono":"3236510681"},{"nombre":"JIMENEZ VELANDIA ALAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3024433396"},{"nombre":"LUGO GONZALEZ MATIAS DAVID","genero":"M","estado":"RETIRADO","telefono":"3241500068"},{"nombre":"OLMOS ECHEVERRY JEILIS THALIANA","genero":"F","estado":"MATRICULADO","telefono":"3007505329"},{"nombre":"PADILLA HERNANDEZ YEILISMAR DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3042097469"},{"nombre":"PEREZ CORENA CARLOS ALFONSO","genero":"M","estado":"MATRICULADO","telefono":"3208081125"},{"nombre":"PEREZ PEREZ VIANIS LUCIA","genero":"F","estado":"MATRICULADO","telefono":"3234856176"},{"nombre":"PUERTA PADILLA RAFAEL GREGORIO","genero":"M","estado":"MATRICULADO","telefono":"3023125659"},{"nombre":"SANDOVAL MARTINEZ MARISOL","genero":"F","estado":"MATRICULADO","telefono":"3127515158"}]},"Septimo B":{"tutor":"JOSEFA MARIA GARCIA MARTINEZ","estudiantes":[{"nombre":"AMAYA SABALA SANTIAGO ANDRES","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"AVILA TINOCO MARIANN SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3014316"},{"nombre":"BADEL HERNANDEZ ACXELL ALFONSO","genero":"M","estado":"MATRICULADO","telefono":"3004813137"},{"nombre":"BARROS OLIVO LEONARDO ANDRES","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"BRIEVA RIOS LAUREN CAMILA","genero":"F","estado":"MATRICULADO","telefono":"3003025175"},{"nombre":"CARRERA GONZALEZ MIGLEIDYSMAR ANGELINA","genero":"F","estado":"MATRICULADO","telefono":"3145686856"},{"nombre":"CERRA LARA DILAN ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"3017655057"},{"nombre":"GAMEZ PADILLA KAROL DANIELA","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"GOMEZ CORONEL LUISA FERNANDA","genero":"F","estado":"MATRICULADO","telefono":"3007173686"},{"nombre":"HENRIQUEZ PINO SYRIANA NIKHOLL","genero":"F","estado":"MATRICULADO","telefono":"3004722968"},{"nombre":"HERRERA DE LA HOZ THIAGO ANDRES","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"MOLINA NUÑEZ GEOVANNI RAFAEL","genero":"M","estado":"MATRICULADO","telefono":"3044953673"},{"nombre":"OSORIO DIAZ DILAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3128767066"},{"nombre":"PADILLA HERNANDEZ ENMANUEL EDUARDO","genero":"M","estado":"MATRICULADO","telefono":"3042097469"},{"nombre":"PARRA CAMACHO MATIAS MAURICIO","genero":"M","estado":"MATRICULADO","telefono":"3024527791"},{"nombre":"PARRA CAMACHO WILLIAMS JOSEP","genero":"M","estado":"MATRICULADO","telefono":"3024527791"},{"nombre":"PRENTT PAJARO JEYCON STIVEN","genero":"M","estado":"MATRICULADO","telefono":"3006658158"},{"nombre":"QUINTERO ORTIZ NELSON JUNIOR","genero":"M","estado":"MATRICULADO","telefono":"3116027128"},{"nombre":"RIPOLL SUAREZ JOSUE DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3015695245"},{"nombre":"RONDON VEGA ENMMANUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"VARGAS ARZUZA WENDY SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3046785492"},{"nombre":"VASQUEZ TABOADA JEREMIS YUSED","genero":"M","estado":"MATRICULADO","telefono":"3042633208"},{"nombre":"ZAPATA PERTUZ NEYDYS ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3004557841"}]},"Septimo C":{"tutor":"ISIDORA DE MOYA MARTÍNEZ","estudiantes":[{"nombre":"AHUMADA OJEDA ELKIN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3042096811"},{"nombre":"ALBA TOLEDO JARED JOSE","genero":"M","estado":"MATRICULADO","telefono":"3107145805"},{"nombre":"ALTAMAR SANCHEZ SCARLETT","genero":"F","estado":"MATRICULADO","telefono":"3042247178"},{"nombre":"AULAR MEZA JEREMIAS DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3239040325"},{"nombre":"AVILA FADUL ANDRY ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3022941805"},{"nombre":"BELLO DE LA HOZ YERINETH CAMILA","genero":"F","estado":"MATRICULADO","telefono":"3044928269"},{"nombre":"CARO SANCHEZ HELLEN SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3024023870"},{"nombre":"ESTRADA PARDO ZHARICK JOHANA","genero":"F","estado":"MATRICULADO","telefono":"3027127253"},{"nombre":"GARCIA RANGEL DANNA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3143311821"},{"nombre":"IBAÑEZ VALLEJO VALERYN MICHEL","genero":"F","estado":"MATRICULADO","telefono":"3122700291"},{"nombre":"MORALES JULIO CAMILO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3112004867"},{"nombre":"PEREZ PORTACIO MATIAS ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3137979568"},{"nombre":"PICALUA TOSCANO KEIDYS MICHELL","genero":"F","estado":"MATRICULADO","telefono":"3022140210"},{"nombre":"QUIROZ MARENCO EMMANUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3023551309"},{"nombre":"RODRIGUEZ ARRIETA HEIDIS ELAYS","genero":"F","estado":"MATRICULADO","telefono":"3207620638"},{"nombre":"RODRIGUEZ DE LOS REYES HILLARY","genero":"F","estado":"MATRICULADO","telefono":"3137262799"},{"nombre":"SALAZAR RAMIREZ VALERIN MICHELL","genero":"F","estado":"MATRICULADO","telefono":"3145607827"},{"nombre":"SOTO ARIZA ANGELIE SOFIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"SOTO DE LA HOZ ALBERTO MARIO","genero":"M","estado":"MATRICULADO","telefono":"3145482902"},{"nombre":"SUAREZ GELIZ LAURA DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3008366474"},{"nombre":"VILLA RODRIGUEZ ALEX DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3145820281"},{"nombre":"WALTEROS GONZALEZ JONATHAN JESUS","genero":"M","estado":"MATRICULADO","telefono":"3013810427"},{"nombre":"AREVALO MORALES LUISA FERNANDA","genero":"F","estado":"RETIRADO","telefono":"3243402383"},{"nombre":"FLOREZ MEDINA JUAN DAVID","genero":"M","estado":"RETIRADO","telefono":""},{"nombre":"GONZALEZ  LUISA ALBERTO","genero":"M","estado":"MATRICULADO","telefono":""}]},"Sexto A":{"tutor":"YULI ANGÉLICA BOHÓRQUEZ LAMBRAÑO","estudiantes":[{"nombre":"AHUMADA MADRID MATIAS ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3054413266"},{"nombre":"ANAYA GALVEZ ISAD FERNANDO","genero":"M","estado":"MATRICULADO","telefono":"3008142584"},{"nombre":"BORRERO MORELO CESAR DE JESUS","genero":"M","estado":"MATRICULADO","telefono":"3003759863"},{"nombre":"CUETO VILLEGAS JOSTIN ANDRES","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"DE LAS SALAS NARVAEZ MELANY","genero":"F","estado":"MATRICULADO","telefono":"3108100840"},{"nombre":"EVERTZ MELENDEZ DHARIANA SHAROT","genero":"F","estado":"MATRICULADO","telefono":"3205706776"},{"nombre":"FAJARDO CARDOZO LAUREN PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3012160174"},{"nombre":"GUEVARA MEJIA MATIAS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3008661623"},{"nombre":"HERRERA BALLESTAS JESUS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3005105013"},{"nombre":"IBAÑEZ VALLEJO DANIELA MICHEL","genero":"F","estado":"MATRICULADO","telefono":"3122700291"},{"nombre":"JURADO RUIZ SOFIA VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3042339238"},{"nombre":"MACHADO PEREA NAYELIS EUCARIS","genero":"F","estado":"MATRICULADO","telefono":"3026482709"},{"nombre":"MEZA SUAREZ VALERITH ANDREINA","genero":"F","estado":"MATRICULADO","telefono":"3025433440"},{"nombre":"NOGUERA CARRASCO YANNELYS SARAI","genero":"F","estado":"MATRICULADO","telefono":"3015742873"},{"nombre":"MORILLO GRANADILLO ANA CAMILA","genero":"F","estado":"MATRICULADO","telefono":"3145043900"},{"nombre":"OLIVEROS HERRERA SANTIAGO DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"ORTIZ ARZUZA DULCE MARIA","genero":"F","estado":"MATRICULADO","telefono":"3116072978"},{"nombre":"PEREZ RODRIGUEZ RICARDO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3242426233"},{"nombre":"ROBLES BALLESTA ENEYEN SELENE","genero":"F","estado":"MATRICULADO","telefono":"3246050514"},{"nombre":"RODAS CALVO SAMUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3113146989"},{"nombre":"RUAS BORRERO JOSE ANTONIO","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"RUA CARABALLO VALERIN JUSITH","genero":"F","estado":"MATRICULADO","telefono":"3008358531"},{"nombre":"VASQUEZ CADRAZCO GABRIELA NICOLL","genero":"F","estado":"MATRICULADO","telefono":"3012234611"},{"nombre":"WALTEROS GONZALEZ GENESIS DEL CARMEN","genero":"F","estado":"MATRICULADO","telefono":"3116851137"},{"nombre":"ZAMBRANO SUAREZ DIEGO ALEXANDER","genero":"M","estado":"RETIRADO","telefono":"3043730560"},{"nombre":"ZUÑIGA COMEZ SAMIR ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":"3015576669"},{"nombre":"DE ORO PEREIRA JESUS ELIAS","genero":"M","estado":"RETIRADO","telefono":"3233695642"},{"nombre":"PATERNINA QUINTERO SHAIRA ANDREA","genero":"F","estado":"RETIRADO","telefono":"3002541102"}]},"Sexto B":{"tutor":"ESPERANZA DE JESUS ORTEGA CANTILLO","estudiantes":[{"nombre":"AMOR PEREZ ISABELLA","genero":"F","estado":"MATRICULADO","telefono":"3004461518"},{"nombre":"BANQUEZ MARRUGO CHARLIS ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3013502411"},{"nombre":"BUSTAMANTE PALENCIA MARIA ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3145456268"},{"nombre":"CALDERON LOPEZ DILAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"6053650996"},{"nombre":"CASTILLEJO ROCHA LUIS ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"3046029316"},{"nombre":"CASTRO MIELES SALOME SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3042043111"},{"nombre":"CHAVIAO HEMER ELIAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3014301233"},{"nombre":"COMENDADOR UTRIA ESTEBAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3103638249"},{"nombre":"CORRALES MORALES ANGEL JOSEPH","genero":"M","estado":"MATRICULADO","telefono":"3045621505"},{"nombre":"DITTA SANTIAGO SAMUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3017502214"},{"nombre":"GONZALEZ SARMIENTO DILAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3003325554"},{"nombre":"HERRERA OBESO JHON DEINER","genero":"M","estado":"MATRICULADO","telefono":"3008131870"},{"nombre":"MARTINEZ ZAMBRANO LUIS GABRIEL","genero":"M","estado":"MATRICULADO","telefono":"3163551175"},{"nombre":"MENDOZA MALLARINO EMANUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3192389720"},{"nombre":"MERCADO SERRANO DEYKOL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3003175425"},{"nombre":"PALLARES GONZALEZ SANTIAGO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3430813"},{"nombre":"POLO PEREZ JEYVER JOSE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"RIVERA MUÑIZ CHERZY NICOL","genero":"F","estado":"MATRICULADO","telefono":"3013644462"},{"nombre":"ROMERO LEON RANDY JOSE","genero":"M","estado":"MATRICULADO","telefono":"3045230139"},{"nombre":"SERPA LOPEZ ALEXANDER JUNIOR","genero":"M","estado":"MATRICULADO","telefono":"3008683029"},{"nombre":"SOTO BARRANCO YULL BREIKEL","genero":"F","estado":"MATRICULADO","telefono":"3016406696"},{"nombre":"TORRES DIAZ MIGUEL ANGEL","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"UTRIA BRICEÑO JEIDERSON ADRIAN","genero":"M","estado":"MATRICULADO","telefono":"3146581471"},{"nombre":"UZCANGA APONTE SEBASTIAN ALEJANDRO","genero":"M","estado":"MATRICULADO","telefono":"3146925332"},{"nombre":"VASQUEZ SANCHEZ JAVIER ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"3004732393"}]},"Sexto C":{"tutor":"STEVENSON LIONS LAGUNA","estudiantes":[{"nombre":"AMAYA PACHON JOSUE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"CARRERA GONZALEZ MAIKEL JOSUE","genero":"M","estado":"MATRICULADO","telefono":"3145686856"},{"nombre":"CASTRO CASTRO ANGEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3013329364"},{"nombre":"COBA URIBE ADRIAN ALBERTO","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"DOMINGUEZ SANDOVAL EDUARDO ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3007768972"},{"nombre":"FLOREZ BOLIVAR RANDY ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3022793052"},{"nombre":"FLOREZ CAMARGO DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3042118912"},{"nombre":"GARCIA PEREZ JEYKO","genero":"M","estado":"MATRICULADO","telefono":"3105826684"},{"nombre":"GOMEZ MENESES JOSUE DAVID","genero":"M","estado":"MATRICULADO","telefono":"3235312372"},{"nombre":"GUTIERREZ CASTRO KLEBERSON MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3017464765"},{"nombre":"HERNANDEZ PADILLA KEIRIS ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3218844757"},{"nombre":"LASCARRO PEÑA DANNA ALEJANDRA","genero":"F","estado":"MATRICULADO","telefono":"3023290180"},{"nombre":"MERCADO VENTURA MAITE MILETH","genero":"F","estado":"MATRICULADO","telefono":"3108986145"},{"nombre":"MERLANO GUTIERREZ FRED DE JESUS","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"NAVARRO LARA ANTHONELLA","genero":"F","estado":"MATRICULADO","telefono":"3052847876"},{"nombre":"ORTIZ BARRAZA YAIRETH","genero":"F","estado":"MATRICULADO","telefono":"3016226245"},{"nombre":"OTERO MARRUGO DEMIAN","genero":"M","estado":"MATRICULADO","telefono":"3045923500"},{"nombre":"PEREZ HENRIQUEZ SAID ESTEBAN","genero":"M","estado":"MATRICULADO","telefono":"3017091020"},{"nombre":"PIÑA GARCIA DARWIN STIVEN","genero":"M","estado":"MATRICULADO","telefono":"3023095970"},{"nombre":"PONCE RAMIREZ JESUS DAVID","genero":"M","estado":"MATRICULADO","telefono":"3045218007"},{"nombre":"RIOS BERDUGO KIANIS MARIA","genero":"F","estado":"MATRICULADO","telefono":"3015833186"},{"nombre":"RUIZ TORO ALFREDO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3016752019"},{"nombre":"SANDOVAL MARTINEZ JUAN DIEGO","genero":"M","estado":"MATRICULADO","telefono":"3127515158"},{"nombre":"SARMIENTO LARA SANTIAGO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3023914981"},{"nombre":"SOLANO CARRACEDO SARID SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3215929360"},{"nombre":"HERNANDEZ DONADO  RICHAR JOSE","genero":"M","estado":"RETIRADO","telefono":""}]},"Tercero A":{"tutor":"LOURDES SANCHEZ LOZANO","estudiantes":[{"nombre":"AGUILAR ACEVEDO DYLAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3004041747"},{"nombre":"ALDANA MEJIA JESUS ADRIAN","genero":"M","estado":"MATRICULADO","telefono":"3015310038"},{"nombre":"ALTUVE GONZALEZ ESTEFANI NAYARITH","genero":"F","estado":"MATRICULADO","telefono":"3014721242"},{"nombre":"ALVERNIA OSPINO LIAM JOSE","genero":"M","estado":"MATRICULADO","telefono":"3016380585"},{"nombre":"ARRIETA PLAZA MARIANNA ESMIL","genero":"F","estado":"MATRICULADO","telefono":"3007250930"},{"nombre":"ARRIETA VASQUEZ JHOSER DAVID","genero":"M","estado":"MATRICULADO","telefono":"3207620638"},{"nombre":"ARROYO PACHECO JOSE MIGUEL","genero":"M","estado":"MATRICULADO","telefono":"3002751002"},{"nombre":"AULAR MEZA ELIAS DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3026110951"},{"nombre":"BARRAZA VASQUEZ SALOME","genero":"F","estado":"MATRICULADO","telefono":"3023864488"},{"nombre":"BARRERA BEITAR LUCIANA VANESSA","genero":"F","estado":"MATRICULADO","telefono":"3117431442"},{"nombre":"BARROS ARRIETA SANTIAGO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3042075277"},{"nombre":"CAMARGO GOMEZ LUCIA","genero":"F","estado":"MATRICULADO","telefono":"3004232440"},{"nombre":"CHAMORRO RAMOS CHISTOPHER ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"3122056866"},{"nombre":"ESCOBAR CELIN LUISA FERNANDA","genero":"F","estado":"MATRICULADO","telefono":"3152505457"},{"nombre":"GARCIA ARAGON ALEXANDER SAID","genero":"M","estado":"MATRICULADO","telefono":"3007448623"},{"nombre":"GARCIA NAVA SOFIA NAEYOMI","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"GONZALEZ CASTILLO THAELL ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"3107141203"},{"nombre":"GRANADILLO CHIRINOS JESUALBER TOMAS","genero":"F","estado":"MATRICULADO","telefono":"3015332819"},{"nombre":"GUERRERO JARAMILLO ANGIE DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3172497752"},{"nombre":"INSIGNARES REYES NIKOLL ANGELIN","genero":"F","estado":"MATRICULADO","telefono":"3147526393"},{"nombre":"LORA ESCALANTE LICEYDIS PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3103598194"},{"nombre":"MACHADO PEREA RONAL MATEO","genero":"M","estado":"MATRICULADO","telefono":"3026482709"},{"nombre":"MERCADO BORRERO VALERY SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3018335091"},{"nombre":"MIRANDA TORRES VALERY SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3136274225"},{"nombre":"MOLINA ROMERO CHRISTOPHER DAVID","genero":"M","estado":"MATRICULADO","telefono":"3023565902"},{"nombre":"OROZCO NARVAEZ JUAN JOSE","genero":"M","estado":"MATRICULADO","telefono":"3017403280"},{"nombre":"PEREZ UBARNE JORGE LUIS","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"POSSO ALVEAR YANNDER YORHANN","genero":"M","estado":"MATRICULADO","telefono":"3004181047"},{"nombre":"RAMIREZ SANTIAGO ISABEL SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3017502214"},{"nombre":"REALES MOGOLLON ABRAHAN ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3244826343"},{"nombre":"ROBLES BALLESTA EDIN JOSE","genero":"M","estado":"MATRICULADO","telefono":"3246050514"},{"nombre":"ROCHA PADILLA FARI SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3128827708"},{"nombre":"SANDOVAL FLOREZ MARIA PAULA","genero":"F","estado":"MATRICULADO","telefono":"3245006691"},{"nombre":"SOTO BARRANCO JOSUE DAVID","genero":"M","estado":"MATRICULADO","telefono":"3244335986"},{"nombre":"UTRIA ALVAREZ ISAIAS JOSE","genero":"M","estado":"MATRICULADO","telefono":"3012205758"},{"nombre":"UTRIA JIMENEZ SNAIDEN DAVID","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"VELEZ TAPIA LUCIA ISABEL","genero":"F","estado":"MATRICULADO","telefono":"3238021271"},{"nombre":"VESGA LOZANO FERNANDO","genero":"M","estado":"MATRICULADO","telefono":"3242507155"},{"nombre":"VILORIA ROMAN LIAM JOSE","genero":"M","estado":"MATRICULADO","telefono":"3003091303"},{"nombre":"VISBAL LEON AARON DAVID","genero":"M","estado":"MATRICULADO","telefono":"3214386311"}]},"Transicion A":{"tutor":"MONICA FRANCO MONTENEGRO","estudiantes":[{"nombre":"ARENAS RUIZ ARIANIS SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3244415177"},{"nombre":"BARCENAS MONTERO BREINER JOSE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"BARRIOS REYES JUSTIN JAVIER","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"DONADO AGUILAR MARIELIS SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3015947944"},{"nombre":"FLOREZ BOLIVAR MATEO FABIAN","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"GAMARRA LARA SALOME","genero":"F","estado":"MATRICULADO","telefono":"3008333147"},{"nombre":"GOMEZ MARTINEZ ABRAHAM","genero":"M","estado":"MATRICULADO","telefono":"3004041789"},{"nombre":"IBARRA CASTILLO ALANNA SOFIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"JIMENEZ TORRES JEICOB JOSE","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"JULIO BELTRAN ORIANA SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3011275273"},{"nombre":"OLIVERO CHALARCA CHRISTOPHER ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"3216315572"},{"nombre":"POLO PEREZ SOFIA NAYERLIN","genero":"F","estado":"MATRICULADO","telefono":"3004489785"},{"nombre":"RAMIREZ PEDROZA MARIEL SHAYITH","genero":"F","estado":"MATRICULADO","telefono":"3015157422"},{"nombre":"REALES JULIO NASLY MILENA","genero":"F","estado":"MATRICULADO","telefono":"3046598340"},{"nombre":"RODRIGUEZ  LIAN MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3006784708"},{"nombre":"ROLONG OSPINO ORIANA LUCIA","genero":"F","estado":"MATRICULADO","telefono":"3044965495"},{"nombre":"ROMERO CRISTANCHO EYLISH SALOME","genero":"F","estado":"MATRICULADO","telefono":"3002805576"},{"nombre":"RUIZ MARAÑON JOSE DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3122056866"},{"nombre":"ZAMBRANO MORALES YULIANIS SOFIA","genero":"F","estado":"MATRICULADO","telefono":"3021193628"},{"nombre":"CASTRO MEDINA GLEIDYMAR VALENTINA","genero":"F","estado":"RETIRADO","telefono":"3025681396"},{"nombre":"PUSHAINA URINA LIAM ANDRES","genero":"M","estado":"RETIRADO","telefono":""}]},"Undecimo A":{"tutor":"MONICA MERCEDES BOSCAN MIELES","estudiantes":[{"nombre":"ARRIETA ANILLO VALERYA PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3013640280"},{"nombre":"BOLAÑO SANCHEZ JULIAN ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"3014046338"},{"nombre":"CADENA SARMIENTO ALEX ORLANDO","genero":"M","estado":"MATRICULADO","telefono":"3004221336"},{"nombre":"CUETO GOMEZ KATIN ALEXANDRA","genero":"F","estado":"MATRICULADO","telefono":"3023509751"},{"nombre":"FLOREZ MOLINARES JESUS MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3246098312"},{"nombre":"GARCIA CARDONA ABIGAIL ESTER","genero":"F","estado":"MATRICULADO","telefono":"3017679789"},{"nombre":"GONZALEZ ALONSO VALENTINA PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3012092278"},{"nombre":"JIMENEZ VELANDIA MATEO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3024433396"},{"nombre":"JURADO AYCARDI NATHALY JOHANA","genero":"F","estado":"MATRICULADO","telefono":"3003515572"},{"nombre":"LLAMAS GONZALEZ YAIBER LEON","genero":"M","estado":"MATRICULADO","telefono":"3015511961"},{"nombre":"MALAGON GAMBOA HAMELLI","genero":"F","estado":"MATRICULADO","telefono":"3017904276"},{"nombre":"MARRIAGA MONTENEGRO LUIS MARIO","genero":"M","estado":"MATRICULADO","telefono":"3016935403"},{"nombre":"MARRIAGA OSORIO YURI MARCELA","genero":"F","estado":"MATRICULADO","telefono":"3147897367"},{"nombre":"MEJIA JIMENEZ LUIS ANGEL","genero":"M","estado":"MATRICULADO","telefono":"3215963015"},{"nombre":"MENDOZA PEREZ ALESKA MARIA","genero":"F","estado":"MATRICULADO","telefono":"3128767277"},{"nombre":"MONTERO CAMPO SARA PATRICK","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"MORENO SANDOVAL RUBEN DARIO","genero":"M","estado":"MATRICULADO","telefono":"3016945428"},{"nombre":"PALOMO BASSA FRANKLIN JESUS","genero":"M","estado":"MATRICULADO","telefono":"3017607762"},{"nombre":"RESTREPO GERALDINO CAMILO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3187251895"},{"nombre":"REYES DE LA HOZ CARLOS ESTEBAN","genero":"M","estado":"MATRICULADO","telefono":"3207263507"},{"nombre":"RODRIGUEZ MIRANDA CARLOS MANUEL","genero":"M","estado":"MATRICULADO","telefono":"3126133479"},{"nombre":"RUIZ MARRIAGA SIANNY LUZ","genero":"F","estado":"MATRICULADO","telefono":"3023999160"},{"nombre":"SAEZ ZUBIRIA MARIA DE LOS ANGELES","genero":"F","estado":"MATRICULADO","telefono":"3004494152"},{"nombre":"SARMIENTO NIETO SANTIAGO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3042040411"},{"nombre":"VALDES SALINAS KARELYS JOHANA","genero":"F","estado":"MATRICULADO","telefono":"3135984300"}]},"Undecimo B":{"tutor":"JUAN CARLOS CASAS JACOME","estudiantes":[{"nombre":"ACOSTA ORTIZ CAMILO ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3004178747"},{"nombre":"ARIZA MARTINEZ SHAIRA YARLETH","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"BAYONA CARMONA JUAN DAVID","genero":"M","estado":"MATRICULADO","telefono":"3107059968"},{"nombre":"BERRIO FERNANDEZ FELIX JOSE","genero":"M","estado":"MATRICULADO","telefono":"3005412559"},{"nombre":"CAMARGO GARCIA JORDAN","genero":"M","estado":"MATRICULADO","telefono":"3043883203"},{"nombre":"CANTILLO TORRES SHERYL ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3003983244"},{"nombre":"CUETER ISSA NAHIARA ISABELA","genero":"F","estado":"MATRICULADO","telefono":"3018099173"},{"nombre":"FERREIRA DE AGUAS JHOJANES","genero":"M","estado":"MATRICULADO","telefono":"3108179758"},{"nombre":"FLORES MALDONADO JOSE DAVID","genero":"M","estado":"MATRICULADO","telefono":"3116710445"},{"nombre":"FONTALVO PAEZ ANGELA ISABEL","genero":"F","estado":"MATRICULADO","telefono":"3136275438"},{"nombre":"FONTALVO TORRES CAROLINA ISABEL","genero":"F","estado":"MATRICULADO","telefono":"3107179550"},{"nombre":"GARCIA ARROYO ALAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3003878372"},{"nombre":"GARCIA HOYOS GENESIS JULIETH","genero":"F","estado":"MATRICULADO","telefono":"3016448818"},{"nombre":"HERNANDEZ OCHOA VALENTINA","genero":"F","estado":"MATRICULADO","telefono":"3177326703"},{"nombre":"HURTADO GARCIA ANDRES FELIPE","genero":"M","estado":"MATRICULADO","telefono":"3157784856"},{"nombre":"NIEBLES MARTINEZ DIANA STEPHANIA","genero":"F","estado":"MATRICULADO","telefono":"3015192453"},{"nombre":"PEREZ PORTACIO JOSE LUIS","genero":"M","estado":"MATRICULADO","telefono":"3007281183"},{"nombre":"QUINTERO POLO MATIAS ESTEBAN","genero":"M","estado":"MATRICULADO","telefono":"3045921126"},{"nombre":"RIVAS HERRERA ALVARO JAVIER","genero":"M","estado":"MATRICULADO","telefono":"3014730489"},{"nombre":"RIVERA SANTOS YONNEY DAVID","genero":"M","estado":"MATRICULADO","telefono":"3024218866"},{"nombre":"RODRIGUEZ COMENDADOR JHOSEPH DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3002107179"},{"nombre":"ROSADO MENDOZA GRACIELA MARIA","genero":"F","estado":"MATRICULADO","telefono":""},{"nombre":"VASQUEZ RUIZ EMANUEL ARMANDO","genero":"M","estado":"MATRICULADO","telefono":"3017465132"},{"nombre":"VILLALOBOS CRUZ JUAN DIEGO","genero":"M","estado":"MATRICULADO","telefono":"3156233929"}]},"Undecimo C":{"tutor":"MIGUEL ANTONIO AMELL MENDOZA","estudiantes":[{"nombre":"ALEAN NAVARRO SAMUEL DAVID","genero":"M","estado":"MATRICULADO","telefono":"3116683529"},{"nombre":"ALZATE MARCELES YISLETH JOHANA","genero":"F","estado":"MATRICULADO","telefono":"3136343223"},{"nombre":"ARAGON CRISTANCHO EMIL JOSUE","genero":"M","estado":"MATRICULADO","telefono":"3177588245"},{"nombre":"ARRIETA RIQUETT JUAN FRANCISCO","genero":"M","estado":"MATRICULADO","telefono":"3157755093"},{"nombre":"AYOLA FLOREZ JUAN CAMILO","genero":"M","estado":"MATRICULADO","telefono":"3117027540"},{"nombre":"CABARCAS BARRIOS PAULA ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3651738"},{"nombre":"CANTILLO ESCORCIA DANIA MARIA","genero":"F","estado":"MATRICULADO","telefono":"3024149825"},{"nombre":"DE LA ROSA FLOREZ SANTIAGO ENRIQUE","genero":"M","estado":"MATRICULADO","telefono":"3012685864"},{"nombre":"DE VIVO OJEDA ANGIE PAOLA","genero":"F","estado":"MATRICULADO","telefono":"3004437007"},{"nombre":"ESCAMILLA RESTREPO ALHAN ALEXANDER","genero":"M","estado":"MATRICULADO","telefono":"3007572254"},{"nombre":"GOMEZ BARRIOS RONALD ELIAS","genero":"M","estado":"MATRICULADO","telefono":"3045486482"},{"nombre":"GONZALEZ BULA JUAN CAMILO","genero":"M","estado":"MATRICULADO","telefono":"3016093534"},{"nombre":"GONZALEZ MEJIA IAN ANDRES","genero":"M","estado":"MATRICULADO","telefono":"3137249835"},{"nombre":"HERRERA JIMENEZ JULIAN ALBERTO","genero":"M","estado":"MATRICULADO","telefono":"6053709887"},{"nombre":"LAURENS PEREZ DILAN YHISAD","genero":"M","estado":"MATRICULADO","telefono":"3012949936"},{"nombre":"LUBO BELTRAN LUIS FERNANDO","genero":"M","estado":"MATRICULADO","telefono":"3042081496"},{"nombre":"MOLINA TORRES ROISY JOHANNA","genero":"F","estado":"MATRICULADO","telefono":"3046582067"},{"nombre":"MOSQUERA TERAN ANTONY DAVID","genero":"M","estado":"MATRICULADO","telefono":"3106949837"},{"nombre":"NAVARRO MUÑOZ NICOLL ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3043420578"},{"nombre":"PEREZ PADILLA LEYDIS DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3218844757"},{"nombre":"PRADO MARCIALES SANTIAGO DAVID","genero":"M","estado":"MATRICULADO","telefono":"3002853942"},{"nombre":"ROA AVILA CHRISTIAN JESUS","genero":"M","estado":"MATRICULADO","telefono":""},{"nombre":"ROMERO LEON FAYSHELL ANDREA","genero":"F","estado":"MATRICULADO","telefono":"3117337537"},{"nombre":"ROMERO RIOS VICTORIA DANIELA","genero":"F","estado":"MATRICULADO","telefono":"3045429483"},{"nombre":"RUIZ VELANDIA JOSUA","genero":"M","estado":"MATRICULADO","telefono":"3104022474"},{"nombre":"SALAZAR MARTINEZ YULAYNIS LORIETH","genero":"F","estado":"MATRICULADO","telefono":"3023369760"},{"nombre":"VASQUEZ MADARIAGA ANGEL DANIEL","genero":"M","estado":"MATRICULADO","telefono":"3007684434"},{"nombre":"TRESPALACIOS SIERRA LUIS SAIR","genero":"M","estado":"RETIRADO","telefono":"3116642528"}]}};
const USUARIOS_INICIALES = {"admin":{"password":"admin2026","role":"admin","nombre":"Administrador","grado":null},"fidia.melgarejo":{"password":"fidia2026","role":"docente","nombre":"FIDIA FERNANDEZ MELGAREJO","grado":"Ciclo 25 A"},"ivaneth.alvarado":{"password":"ivaneth2026","role":"docente","nombre":"IVANETH IBAÑEZ ALVARADO","grado":"Cuarto A"},"luzdary.camargo":{"password":"luzdary2026","role":"docente","nombre":"LUZDARY EGEA CAMARGO","grado":"Cuarto B"},"deysi.barrios":{"password":"deysi2026","role":"docente","nombre":"DEYSI SOFIA RUIZ BARRIOS","grado":"Decimo A"},"fanny.cantillo":{"password":"fanny2026","role":"docente","nombre":"FANNY ISABEL AHUMADA CANTILLO","grado":"Decimo B"},"nelda.barandica":{"password":"nelda2026","role":"docente","nombre":"NELDA FONTALVO BARANDICA","grado":"Jardin A"},"juan.castañeda":{"password":"juan2026","role":"docente","nombre":"JUAN ANTONIO CONSTANTE CASTAÑEDA","grado":"Noveno A"},"lotty.peñaloza":{"password":"lotty2026","role":"docente","nombre":"LOTTY JUDITH CAMARGO PEÑALOZA","grado":"Noveno B"},"noelvis.alvarez":{"password":"noelvis2026","role":"docente","nombre":"NOELVIS JOSEFINA MENDEZ ALVAREZ","grado":"Octavo A"},"zulima.carrillo":{"password":"zulima2026","role":"docente","nombre":"ZULIMA LUCIA RODRIGUEZ CARRILLO","grado":"Octavo B"},"luz.cuentas":{"password":"luz2026","role":"docente","nombre":"LUZ ELENA DE LA HOZ CUENTAS","grado":"Primero A"},"nora.rodriguez":{"password":"nora2026","role":"docente","nombre":"NORA GONZALEZ RODRIGUEZ","grado":"Primero B"},"rocio.sabalza":{"password":"rocio2026","role":"docente","nombre":"ROCIO MUÑOZ SABALZA","grado":"Quinto A"},"berenice.borja":{"password":"berenice2026","role":"docente","nombre":"BERENICE MERCADO BORJA","grado":"Quinto B"},"carmen.olivares":{"password":"carmen2026","role":"docente","nombre":"CARMEN AHUMADA OLIVARES","grado":"Segundo A"},"claribel.ojeda":{"password":"claribel2026","role":"docente","nombre":"CLARIBEL CELESTE REALES OJEDA","grado":"Septimo A"},"josefa.martinez":{"password":"josefa2026","role":"docente","nombre":"JOSEFA MARIA GARCIA MARTINEZ","grado":"Septimo B"},"isidora.martínez":{"password":"isidora2026","role":"docente","nombre":"ISIDORA DE MOYA MARTÍNEZ","grado":"Septimo C"},"yuli.lambraño":{"password":"yuli2026","role":"docente","nombre":"YULI ANGÉLICA BOHÓRQUEZ LAMBRAÑO","grado":"Sexto A"},"esperanza.cantillo":{"password":"esperanza2026","role":"docente","nombre":"ESPERANZA DE JESUS ORTEGA CANTILLO","grado":"Sexto B"},"stevenson.laguna":{"password":"stevenson2026","role":"docente","nombre":"STEVENSON LIONS LAGUNA","grado":"Sexto C"},"lourdes.lozano":{"password":"lourdes2026","role":"docente","nombre":"LOURDES SANCHEZ LOZANO","grado":"Tercero A"},"monica.montenegro":{"password":"monica2026","role":"docente","nombre":"MONICA FRANCO MONTENEGRO","grado":"Transicion A"},"monica.mieles":{"password":"monica2026","role":"docente","nombre":"MONICA MERCEDES BOSCAN MIELES","grado":"Undecimo A"},"juan.jacome":{"password":"juan2026","role":"docente","nombre":"JUAN CARLOS CASAS JACOME","grado":"Undecimo B"},"miguel.mendoza":{"password":"miguel2026","role":"docente","nombre":"MIGUEL ANTONIO AMELL MENDOZA","grado":"Undecimo C"},"coordinador":{"password":"coord2026","role":"coordinador","nombre":"Coordinación","grado":null}};
const GRADOS_ORDEN = ["Jardin A", "Transicion A", "Primero A", "Primero B", "Segundo A", "Tercero A", "Cuarto A", "Cuarto B", "Quinto A", "Quinto B", "Sexto A", "Sexto B", "Sexto C", "Septimo A", "Septimo B", "Septimo C", "Octavo A", "Octavo B", "Noveno A", "Noveno B", "Decimo A", "Decimo B", "Undecimo A", "Undecimo B", "Undecimo C", "Ciclo 23 A", "Ciclo 24 A", "Ciclo 25 A"];

const ESTADOS = [
  { key:"MATRICULADO",             label:"Matriculado",          color:"#1e7e34", bg:"#d4edda" },
  { key:"RETIRADO",                label:"Retirado",             color:"#721c24", bg:"#f8d7da" },
  { key:"REPROBADO",               label:"Reprobado",            color:"#856404", bg:"#fff3cd" },
  { key:"PROMOVIDO",               label:"Promovido",            color:"#004085", bg:"#cce5ff" },
  { key:"COMPROMISO CONVIVENCIAL", label:"Comp. Convivencial",   color:"#5a1e6b", bg:"#f3d9fa" },
  { key:"COMPROMISO ACADEMICO",    label:"Comp. Académico",      color:"#7c4a00", bg:"#ffe8cc" },
  { key:"CONTINUIDAD",             label:"Continuidad",          color:"#495057", bg:"#e9ecef" },
];
const SIMAT_ESTADOS = [
  { key:"REGISTRADO",    label:"Registrado",    color:"#1e7e34", bg:"#d4edda" },
  { key:"NO LIBERADO",   label:"No Liberado",   color:"#856404", bg:"#fff3cd" },
  { key:"NO REGISTRADO", label:"No Registrado", color:"#721c24", bg:"#f8d7da" },
];
const SEMANAS=["S1","S2","S3","S4"],DIAS=["L","M","M","J","V"],PERIODOS=["P1","P2","P3","P4"],SABERES=["Saber","Saber Hacer","Ser"],DIAS_LABELS=["Lun","Mar","Mié","Jue","Vie"];
const STORAGE_KEY="inedisav_2026_v5",AUTH_KEY="inedisav_auth_v5",USERS_KEY="inedisav_users_v5",DATA_KEY="inedisav_data_v5";

function loadStorage(key,fallback){try{const r=localStorage.getItem(key);return r?JSON.parse(r):fallback;}catch{return fallback;}}
function saveStorage(key,val){try{localStorage.setItem(key,JSON.stringify(val));}catch{}}
function estadoInfo(k){return ESTADOS.find(e=>e.key===k)||ESTADOS[0];}
function simatInfo(k){return SIMAT_ESTADOS.find(s=>s.key===k)||SIMAT_ESTADOS[2];}
function pct(n,d){return d===0?"0%":(n/d*100).toFixed(1)+"%";}

const th={padding:"8px 10px",fontWeight:700,fontSize:12,color:"#374151",borderRight:"1px solid #e9ecf0",textAlign:"center",whiteSpace:"nowrap"};
const td={padding:"6px 10px",borderRight:"1px solid #f0f4f8",verticalAlign:"middle"};

// ══════════ LOGIN ══════════
function LoginScreen({onLogin}){
  const [user,setUser]=useState(""),[ pass,setPass]=useState(""),[ err,setErr]=useState(""),[ show,setShow]=useState(false);
  function handleLogin(){
    const usuarios=loadStorage(USERS_KEY,USUARIOS_INICIALES);
    const u=usuarios[user.trim().toLowerCase()];
    if(!u||u.password!==pass){setErr("Usuario o contraseña incorrectos");return;}
    saveStorage(AUTH_KEY,{username:user.trim().toLowerCase(),role:u.role,nombre:u.nombre,grado:u.grado});
    onLogin({username:user.trim().toLowerCase(),role:u.role,nombre:u.nombre,grado:u.grado});
  }
  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0d47a1,#1565c0)",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"40px 36px",width:340,boxShadow:"0 8px 32px rgba(0,0,0,.25)"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{fontSize:36}}>🏫</div>
          <div style={{fontWeight:800,fontSize:18,color:"#1e3a5f",marginTop:8}}>INEDISAV</div>
          <div style={{fontSize:12,color:"#6b7280",marginTop:2}}>IED Sarid Arteta de Vásquez · 2026</div>
        </div>
        <div style={{marginBottom:14}}>
          <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Usuario</label>
          <input value={user} onChange={e=>setUser(e.target.value)} placeholder="ej: admin"
            onKeyDown={e=>e.key==="Enter"&&handleLogin()}
            style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:14,boxSizing:"border-box"}}/>
        </div>
        <div style={{marginBottom:6}}>
          <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Contraseña</label>
          <div style={{position:"relative"}}>
            <input value={pass} onChange={e=>setPass(e.target.value)} type={show?"text":"password"} placeholder="••••••••"
              onKeyDown={e=>e.key==="Enter"&&handleLogin()}
              style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:14,boxSizing:"border-box"}}/>
            <button onClick={()=>setShow(!show)} style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",fontSize:16,color:"#6b7280"}}>{show?"🙈":"👁️"}</button>
          </div>
        </div>
        {err&&<div style={{color:"#dc3545",fontSize:12,marginBottom:8,textAlign:"center"}}>{err}</div>}
        <button onClick={handleLogin} style={{width:"100%",padding:"11px",borderRadius:8,background:"#1565c0",color:"#fff",fontWeight:700,fontSize:15,border:"none",cursor:"pointer",marginTop:10}}>Ingresar</button>
        <div style={{marginTop:16,fontSize:11,color:"#9ca3af",textAlign:"center"}}>Admin: <b>admin</b> / <b>admin2026</b></div>
      </div>
    </div>
  );
}

// ══════════ CREAR USUARIO (sub-componente) ══════════
function CrearUsuarioForm({usuarios,setUsuarios,datosEscuela}){
  const [show,setShow]=useState(false);
  const [form,setForm]=useState({username:"",password:"",nombre:"",role:"docente",grado:""});
  const [msg,setMsg]=useState("");
  const grados=GRADOS_ORDEN.filter(g=>datosEscuela[g]);
  function crear(){
    if(!form.username.trim()){setMsg("⚠️ Usuario obligatorio");return;}
    if(!form.password.trim()){setMsg("⚠️ Contraseña obligatoria");return;}
    if(!form.nombre.trim()){setMsg("⚠️ Nombre obligatorio");return;}
    if(form.role==="docente"&&!form.grado){setMsg("⚠️ Selecciona un grado");return;}
    const key=form.username.trim().toLowerCase().replace(/\s/g,"");
    if(usuarios[key]){setMsg("⚠️ Usuario ya existe");return;}
    const updated={...usuarios,[key]:{password:form.password.trim(),role:form.role,nombre:form.nombre.trim().toUpperCase(),grado:form.role==="docente"?form.grado:null}};
    setUsuarios(updated);saveStorage(USERS_KEY,updated);
    setMsg("✅ Usuario creado: "+key);
    setForm({username:"",password:"",nombre:"",role:"docente",grado:""});
    setTimeout(()=>setMsg(""),3000);
  }
  return(
    <div style={{marginBottom:16}}>
      <button onClick={()=>setShow(!show)} style={{padding:"8px 18px",borderRadius:8,background:"#1565c0",color:"#fff",fontWeight:700,fontSize:13,border:"none",cursor:"pointer",marginBottom:show?12:0}}>➕ Crear nuevo usuario</button>
      {show&&(
        <div style={{background:"#f0f4ff",border:"1.5px solid #1565c030",borderRadius:10,padding:16}}>
          {msg&&<div style={{padding:"7px 12px",borderRadius:7,marginBottom:10,fontSize:12,fontWeight:600,background:msg.startsWith("✅")?"#d4edda":"#fff3cd",color:msg.startsWith("✅")?"#1e7e34":"#856404"}}>{msg}</div>}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:3}}>Usuario *</label>
              <input value={form.username} onChange={e=>setForm({...form,username:e.target.value.toLowerCase()})} placeholder="ej: maria.perez"
                style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"1.5px solid #d1d5db",fontSize:12,boxSizing:"border-box"}}/></div>
            <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:3}}>Contraseña *</label>
              <input value={form.password} onChange={e=>setForm({...form,password:e.target.value})} placeholder="ej: maria2026"
                style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"1.5px solid #d1d5db",fontSize:12,boxSizing:"border-box"}}/></div>
            <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:3}}>Nombre completo *</label>
              <input value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})} placeholder="ej: MARIA PEREZ"
                style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"1.5px solid #d1d5db",fontSize:12,boxSizing:"border-box"}}/></div>
            <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:3}}>Rol *</label>
              <select value={form.role} onChange={e=>setForm({...form,role:e.target.value,grado:""})}
                style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"1.5px solid #d1d5db",fontSize:12}}>
                <option value="docente">👤 Docente</option>
                <option value="coordinador">📋 Coordinador</option>
                <option value="admin">👑 Administrador</option>
              </select></div>
            {form.role==="docente"&&(
              <div style={{gridColumn:"1/-1"}}><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:3}}>Grado *</label>
                <select value={form.grado} onChange={e=>setForm({...form,grado:e.target.value})}
                  style={{width:"100%",padding:"7px 10px",borderRadius:7,border:"1.5px solid #d1d5db",fontSize:12}}>
                  <option value="">— Selecciona —</option>
                  {grados.map(g=><option key={g} value={g}>{g} — {datosEscuela[g]?.tutor||""}</option>)}
                </select></div>
            )}
          </div>
          <button onClick={crear} style={{padding:"8px 20px",borderRadius:7,background:"#1e7e34",color:"#fff",fontWeight:700,fontSize:13,border:"none",cursor:"pointer"}}>✅ Guardar usuario</button>
        </div>
      )}
    </div>
  );
}

// ══════════ EDITAR ESTUDIANTE (sub-componente) ══════════
function EditarEstudiante({est,onSave,onCancel}){
  const [d,setD]=useState({...est});
  return(
    <div style={{background:"#f0f4ff",border:"2px solid #1565c040",padding:12,margin:4,borderRadius:8}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:8}}>
        <div style={{gridColumn:"1/-1"}}><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Apellidos y Nombres</label>
          <input value={d.nombre||""} onChange={e=>setD({...d,nombre:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,boxSizing:"border-box"}}/></div>
        <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Género</label>
          <select value={d.genero||"M"} onChange={e=>setD({...d,genero:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12}}>
            <option value="M">M</option><option value="F">F</option></select></div>
        <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Fecha Nacimiento</label>
          <input type="date" value={d.fechaNac||""} onChange={e=>setD({...d,fechaNac:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,boxSizing:"border-box"}}/></div>
        <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Documento</label>
          <input value={d.documento||""} onChange={e=>setD({...d,documento:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,boxSizing:"border-box"}}/></div>
        <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>PERID</label>
          <input value={d.perid||""} onChange={e=>setD({...d,perid:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,boxSizing:"border-box"}}/></div>
        <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Teléfono</label>
          <input value={d.telefono||""} onChange={e=>setD({...d,telefono:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,boxSizing:"border-box"}}/></div>
        <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Ruta</label>
          <input value={d.ruta||""} onChange={e=>setD({...d,ruta:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,boxSizing:"border-box"}}/></div>
        <div><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Transportador</label>
          <input value={d.transportador||""} onChange={e=>setD({...d,transportador:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,boxSizing:"border-box"}}/></div>
        <div style={{gridColumn:"1/-1"}}><label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:2}}>Estado SIMAT</label>
          <select value={d.simat||"NO REGISTRADO"} onChange={e=>setD({...d,simat:e.target.value})}
            style={{width:"100%",padding:"6px 9px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12}}>
            {SIMAT_ESTADOS.map(s=><option key={s.key} value={s.key}>{s.label}</option>)}
          </select></div>
      </div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={()=>onSave(d)} style={{padding:"6px 16px",borderRadius:6,background:"#1565c0",color:"#fff",fontWeight:700,fontSize:12,border:"none",cursor:"pointer"}}>💾 Guardar</button>
        <button onClick={onCancel} style={{padding:"6px 14px",borderRadius:6,background:"#e9ecef",color:"#374151",fontWeight:600,fontSize:12,border:"none",cursor:"pointer"}}>Cancelar</button>
      </div>
    </div>
  );
}

// ══════════ LISTA ESTUDIANTES EDITABLE ══════════
function ListaEstudiantes({grado,datosEscuela,onDataUpdate}){
  const [editIdx,setEditIdx]=useState(null);
  const [msg,setMsg]=useState("");
  const ests=datosEscuela[grado]?.estudiantes||[];
  function handleSave(data){
    if(!data.nombre?.trim()){setMsg("⚠️ Nombre obligatorio");return;}
    const updated=JSON.parse(JSON.stringify(datosEscuela));
    updated[grado].estudiantes[editIdx]={...data,nombre:data.nombre.trim().toUpperCase()};
    saveStorage(DATA_KEY,updated);onDataUpdate(updated);
    const cur=loadStorage(STORAGE_KEY,{});
    if(!cur[grado])cur[grado]={};if(!cur[grado].simat)cur[grado].simat={};
    cur[grado].simat[editIdx]=data.simat||"NO REGISTRADO";
    saveStorage(STORAGE_KEY,cur);
    setEditIdx(null);setMsg("✅ Actualizado");setTimeout(()=>setMsg(""),2500);
  }
  function handleDelete(i){
    if(!window.confirm("¿Eliminar este estudiante?"))return;
    const updated=JSON.parse(JSON.stringify(datosEscuela));
    updated[grado].estudiantes.splice(i,1);
    saveStorage(DATA_KEY,updated);onDataUpdate(updated);
    const cur=loadStorage(STORAGE_KEY,{});
    const gs=cur[grado]||{};
    ["comisiones","asistencia","notas","entregas","simat"].forEach(sec=>{
      if(!gs[sec])return;
      const ri={};
      Object.entries(gs[sec]).forEach(([idx,val])=>{const n=parseInt(idx);if(n<i)ri[n]=val;else if(n>i)ri[n-1]=val;});
      gs[sec]=ri;
    });
    cur[grado]=gs;saveStorage(STORAGE_KEY,cur);
    if(editIdx===i)setEditIdx(null);
    setMsg("✅ Eliminado");setTimeout(()=>setMsg(""),2500);
  }
  return(
    <div style={{marginTop:20}}>
      <div style={{fontSize:12,fontWeight:700,color:"#1e3a5f",marginBottom:8}}>Estudiantes en {grado} ({ests.length}) — ✏️ editar · 🗑️ eliminar</div>
      {msg&&<div style={{padding:"6px 12px",borderRadius:7,marginBottom:8,fontSize:12,fontWeight:600,background:msg.startsWith("✅")?"#d4edda":"#fff3cd",color:msg.startsWith("✅")?"#1e7e34":"#856404"}}>{msg}</div>}
      <div style={{maxHeight:340,overflowY:"auto",border:"1px solid #e9ecf0",borderRadius:10}}>
        {ests.map((e,i)=>(
          <div key={i}>
            {editIdx===i
              ? <EditarEstudiante est={e} onSave={handleSave} onCancel={()=>setEditIdx(null)}/>
              : <div style={{padding:"7px 12px",fontSize:12,borderBottom:"1px solid #f0f4f8",background:i%2===0?"#fff":"#f9fbff",display:"flex",gap:8,alignItems:"center"}}>
                  <span style={{color:"#9ca3af",minWidth:24}}>{i+1}</span>
                  <span style={{flex:1,fontWeight:500}}>{e.nombre}</span>
                  <span style={{color:"#6b7280",minWidth:16}}>{e.genero}</span>
                  {e.fechaNac&&<span style={{color:"#6b7280",fontSize:11}}>🎂 {e.fechaNac}</span>}
                  {e.ruta&&<span style={{color:"#1565c0",fontSize:11}}>🚌 {e.ruta}</span>}
                  <button onClick={()=>setEditIdx(i)} style={{padding:"3px 10px",borderRadius:5,background:"#e8edf5",border:"none",cursor:"pointer",fontSize:11,color:"#1565c0",fontWeight:600}}>✏️</button>
                  <button onClick={()=>handleDelete(i)} style={{padding:"3px 10px",borderRadius:5,background:"#f8d7da",border:"none",cursor:"pointer",fontSize:11,color:"#721c24",fontWeight:600}}>🗑️</button>
                </div>
            }
          </div>
        ))}
      </div>
    </div>
  );
}

// ══════════ NUEVO ESTUDIANTE ══════════
function NuevoEstudianteTab({datosEscuela,onDataUpdate}){
  const grados=GRADOS_ORDEN.filter(g=>datosEscuela[g]);
  const [form,setForm]=useState({nombre:"",genero:"M",documento:"",perid:"",simatEst:"NO REGISTRADO",telefono:"",grado:grados[0]||"",ruta:"",transportador:"",fechaNac:""});
  const [msg,setMsg]=useState("");
  function agregar(){
    if(!form.nombre.trim()){setMsg("⚠️ Nombre obligatorio");return;}
    if(!form.grado){setMsg("⚠️ Selecciona un grado");return;}
    const updated=JSON.parse(JSON.stringify(datosEscuela));
    updated[form.grado].estudiantes.push({nombre:form.nombre.trim().toUpperCase(),genero:form.genero,documento:form.documento.trim(),perid:form.perid.trim(),simat:form.simatEst||"NO REGISTRADO",estado:"MATRICULADO",telefono:form.telefono.trim(),ruta:form.ruta.trim(),transportador:form.transportador.trim(),fechaNac:form.fechaNac});
    saveStorage(DATA_KEY,updated);onDataUpdate(updated);
    setMsg("✅ Agregado a "+form.grado);
    setForm({nombre:"",genero:"M",documento:"",perid:"",simatEst:"NO REGISTRADO",telefono:"",grado:form.grado,ruta:"",transportador:"",fechaNac:""});
    setTimeout(()=>setMsg(""),3000);
  }
  const F=(label,field,placeholder,type="text")=>(
    <div><label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>{label}</label>
      <input type={type} value={form[field]} onChange={e=>setForm({...form,[field]:e.target.value})} placeholder={placeholder}
        style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,boxSizing:"border-box"}}/></div>
  );
  return(
    <div>
      <div style={{fontSize:13,color:"#6b7280",marginBottom:16}}>Agrega un estudiante. Aparecerá en todos los formatos con estado <b>Matriculado</b>.</div>
      {msg&&<div style={{padding:"8px 12px",borderRadius:8,marginBottom:12,fontSize:13,fontWeight:600,background:msg.startsWith("✅")?"#d4edda":"#fff3cd",color:msg.startsWith("✅")?"#1e7e34":"#856404"}}>{msg}</div>}
      <div style={{display:"grid",gap:12}}>
        {F("Apellidos y Nombres *","nombre","ej: GARCIA LOPEZ JUAN PABLO")}
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          <div><label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Género</label>
            <select value={form.genero} onChange={e=>setForm({...form,genero:e.target.value})}
              style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13}}>
              <option value="M">Masculino (M)</option><option value="F">Femenino (F)</option></select></div>
          {F("Teléfono","telefono","ej: 3001234567")}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          {F("Documento","documento","ej: 1234567890")}
          {F("PERID","perid","ej: 123456")}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          {F("Fecha de Nacimiento","fechaNac","","date")}
          <div><label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Estado SIMAT</label>
            <select value={form.simatEst} onChange={e=>setForm({...form,simatEst:e.target.value})}
              style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13}}>
              {SIMAT_ESTADOS.map(s=><option key={s.key} value={s.key}>{s.label}</option>)}
            </select></div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          {F("Ruta","ruta","ej: Ruta 3")}
          {F("Transportador","transportador","ej: Juan Pérez")}
        </div>
        <div><label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:4}}>Grado *</label>
          <select value={form.grado} onChange={e=>setForm({...form,grado:e.target.value})}
            style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13}}>
            {grados.map(g=><option key={g} value={g}>{g} — {datosEscuela[g]?.tutor||""}</option>)}
          </select></div>
        <button onClick={agregar} style={{padding:"11px",borderRadius:8,background:"#1565c0",color:"#fff",fontWeight:700,fontSize:14,border:"none",cursor:"pointer"}}>➕ Agregar Estudiante</button>
      </div>
      <ListaEstudiantes grado={form.grado} datosEscuela={datosEscuela} onDataUpdate={onDataUpdate}/>
    </div>
  );
}

// ══════════ TIPOS / GRADOS ══════════
function TiposTab({datosEscuela,onDataUpdate}){
  const init=GRADOS_ORDEN.filter(g=>datosEscuela[g]).map(g=>({grado:g,tutor:datosEscuela[g]?.tutor||""}));
  const [filas,setFilas]=useState(init.length>0?init:[{grado:"",tutor:""}]);
  const [msg,setMsg]=useState("");
  function guardar(){
    if(filas.some(f=>!f.grado.trim())){setMsg("⚠️ Hay filas sin grado");return;}
    const updated=JSON.parse(JSON.stringify(datosEscuela));
    filas.forEach(f=>{const g=f.grado.trim();if(!updated[g])updated[g]={tutor:f.tutor.trim(),estudiantes:[]};else updated[g].tutor=f.tutor.trim();});
    saveStorage(DATA_KEY,updated);onDataUpdate(updated);
    setMsg("✅ Guardado");setTimeout(()=>setMsg(""),3000);
  }
  return(
    <div>
      <div style={{fontSize:13,color:"#6b7280",marginBottom:12}}>Gestiona grados y directores. Los cambios aplican en todos los formatos.</div>
      {msg&&<div style={{padding:"8px 12px",borderRadius:8,marginBottom:12,fontSize:13,fontWeight:600,background:msg.startsWith("✅")?"#d4edda":"#fff3cd",color:msg.startsWith("✅")?"#1e7e34":"#856404"}}>{msg}</div>}
      <div style={{border:"1px solid #e9ecf0",borderRadius:10,overflow:"hidden",marginBottom:12}}>
        <div style={{background:"#1565c0",color:"#fff",padding:"8px 14px",display:"grid",gridTemplateColumns:"40px 1fr 1fr 40px",gap:8,fontSize:12,fontWeight:700}}>
          <span>#</span><span>Grado</span><span>Director(a) de Grupo</span><span></span>
        </div>
        <div style={{maxHeight:380,overflowY:"auto"}}>
          {filas.map((f,i)=>(
            <div key={i} style={{display:"grid",gridTemplateColumns:"40px 1fr 1fr 40px",gap:8,padding:"6px 10px",background:i%2===0?"#fff":"#f9fbff",borderBottom:"1px solid #f0f4f8",alignItems:"center"}}>
              <span style={{fontSize:12,color:"#9ca3af",textAlign:"center"}}>{i+1}</span>
              <input value={f.grado} onChange={e=>{const c=[...filas];c[i]={...c[i],grado:e.target.value};setFilas(c);}} placeholder="ej: Primero A"
                style={{padding:"5px 8px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,width:"100%",boxSizing:"border-box"}}/>
              <input value={f.tutor} onChange={e=>{const c=[...filas];c[i]={...c[i],tutor:e.target.value};setFilas(c);}} placeholder="ej: MARIA GARCIA"
                style={{padding:"5px 8px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,width:"100%",boxSizing:"border-box"}}/>
              <button onClick={()=>setFilas(filas.filter((_,idx)=>idx!==i))} style={{background:"#f8d7da",border:"none",borderRadius:6,color:"#721c24",fontWeight:700,cursor:"pointer",fontSize:14,padding:"4px 8px"}}>✕</button>
            </div>
          ))}
        </div>
      </div>
      <div style={{display:"flex",gap:10}}>
        <button onClick={()=>setFilas([...filas,{grado:"",tutor:""}])} style={{padding:"8px 18px",borderRadius:8,background:"#e9ecef",color:"#374151",fontWeight:600,fontSize:13,border:"1.5px solid #d1d5db",cursor:"pointer"}}>➕ Agregar grado</button>
        <button onClick={guardar} style={{padding:"8px 24px",borderRadius:8,background:"#1565c0",color:"#fff",fontWeight:700,fontSize:13,border:"none",cursor:"pointer"}}>💾 Guardar cambios</button>
      </div>
    </div>
  );
}

// ══════════ ADMIN PANEL ══════════
function AdminPanel({onClose,datosEscuela,onDataUpdate}){
  const [subTab,setSubTab]=useState("usuarios");
  const [usuarios,setUsuarios]=useState(()=>loadStorage(USERS_KEY,USUARIOS_INICIALES));
  const [editUser,setEditUser]=useState(null);
  const [newPass,setNewPass]=useState("");
  const [msg,setMsg]=useState("");
  const [simatMsg,setSimatMsg]=useState("");
  const [confirmarBorrar,setConfirmarBorrar]=useState(false);

  function guardarPass(username){
    if(!newPass.trim())return;
    const updated={...usuarios,[username]:{...usuarios[username],password:newPass.trim()}};
    setUsuarios(updated);saveStorage(USERS_KEY,updated);
    setEditUser(null);setNewPass("");setMsg("✅ Contraseña actualizada");setTimeout(()=>setMsg(""),2000);
  }
  function borrarBD(){
    localStorage.clear();saveStorage(DATA_KEY,DATOS_ESCUELA_INICIAL);saveStorage(USERS_KEY,USUARIOS_INICIALES);window.location.reload();
  }
  function handleSIMAT(e){
    const file=e.target.files[0];if(!file)return;setSimatMsg("Procesando...");
    const reader=new FileReader();
    reader.onload=(ev)=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"array"});
        const simat=XLSX.utils.sheet_to_json(wb.Sheets["SIMAT"]||wb.Sheets[wb.SheetNames[0]]);
        const tipos=XLSX.utils.sheet_to_json(wb.Sheets["Tipos"]||wb.Sheets[wb.SheetNames[1]],{header:1});
        const tutorMap={};
        tipos.forEach(row=>{if(row[17]&&row[18])tutorMap[String(row[17]).trim()]=String(row[18]).trim();});
        const newData={};
        simat.forEach(r=>{
          const grado=(r["GRADO"]||r["Grado"]||r["grado"]||"").toString().trim();if(!grado)return;
          if(!newData[grado])newData[grado]={tutor:tutorMap[grado]||"",estudiantes:[]};
          let tel="";try{const t=parseInt(String(r["TELEFONO"]||r["Telefono"]||"").replace(/[^0-9]/g,""));if(t&&t>0)tel=String(t);}catch{}
          const nombre=(r["APELLIDOS_Y_NOMBRES"]||r["Apellidos y Nombres"]||"").toString().trim();
          if(!nombre)return;
          const genero=(r["GENERO"]||r["Genero"]||"").toString().trim().toUpperCase();
          const estadoRaw=(r["COMISION FINAL"]||r["Comision Final"]||"MATRICULADO").toString().trim().toUpperCase();
          const estado=estadoRaw&&estadoRaw!=="NAN"?estadoRaw:"MATRICULADO";
          const simatVal=(r["SIMAT"]||r["Simat"]||"NO REGISTRADO").toString().trim().toUpperCase();
          const ruta=(r["RUTA"]||r["Ruta"]||"").toString().trim();
          const transportador=(r["TRANSPORTADOR"]||r["Transportador"]||"").toString().trim();
          const fechaNac=(r["FECHA_NACIMIENTO"]||r["Fecha Nacimiento"]||r["FECHA NACIMIENTO"]||"").toString().trim();
          newData[grado].estudiantes.push({nombre,genero,estado,simat:simatVal||"NO REGISTRADO",telefono:tel,ruta,transportador,fechaNac});
        });
        saveStorage(DATA_KEY,newData);onDataUpdate(newData);
        setSimatMsg("✅ SIMAT cargado: "+Object.keys(newData).length+" grados, "+simat.length+" estudiantes");
      }catch(err){setSimatMsg("❌ Error: "+err.message);}
    };
    reader.readAsArrayBuffer(file);
  }

  const TABS=[["usuarios","👥 Usuarios"],["estudiantes","➕ Nuevo Estudiante"],["tipos","🏫 Grados y Tutores"],["simat","📂 Importar SIMAT"],["borrar","🗑️ Borrar BD"]];

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"#fff",borderRadius:16,width:"min(720px,96vw)",maxHeight:"88vh",display:"flex",flexDirection:"column",boxShadow:"0 8px 40px rgba(0,0,0,.3)"}}>
        <div style={{background:"#1565c0",color:"#fff",padding:"16px 20px",borderRadius:"16px 16px 0 0",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontWeight:700,fontSize:15}}>⚙️ Panel de Administración</span>
          <button onClick={onClose} style={{background:"none",border:"none",color:"#fff",fontSize:20,cursor:"pointer"}}>✕</button>
        </div>
        <div style={{display:"flex",borderBottom:"2px solid #e3e8f0",padding:"0 20px",overflowX:"auto"}}>
          {TABS.map(([k,l])=>(
            <button key={k} onClick={()=>setSubTab(k)}
              style={{padding:"10px 14px",border:"none",background:"none",cursor:"pointer",fontWeight:subTab===k?700:400,whiteSpace:"nowrap",
                color:subTab===k?"#1565c0":"#6b7280",borderBottom:subTab===k?"3px solid #1565c0":"3px solid transparent",marginBottom:-2,fontSize:13}}>
              {l}
            </button>
          ))}
        </div>
        <div style={{padding:20,overflowY:"auto",flex:1}}>
          {msg&&<div style={{background:"#d4edda",color:"#1e7e34",padding:"8px 12px",borderRadius:8,marginBottom:12,fontSize:13}}>{msg}</div>}

          {subTab==="usuarios"&&(
            <div>
              <CrearUsuarioForm usuarios={usuarios} setUsuarios={setUsuarios} datosEscuela={datosEscuela}/>
              {/* Admin */}
              {[["admin",...[usuarios["admin"]]],...Object.entries(usuarios).filter(([k])=>k!=="admin")].map(([username,u])=>{
                if(!u)return null;
                const icon=u.role==="admin"?"👑":u.role==="coordinador"?"📋":"👤";
                const label=u.role==="admin"?"Administrador":u.role==="coordinador"?"Coordinador":"Docente";
                const bg=u.role==="admin"?"#e8edf5":u.role==="coordinador"?"#f3d9fa":"#fff";
                return(
                  <div key={username} style={{border:"1px solid #e9ecf0",borderRadius:8,padding:"10px 14px",marginBottom:6,display:"flex",alignItems:"center",gap:12,background:bg}}>
                    <span style={{fontSize:18}}>{icon}</span>
                    <div style={{flex:1}}>
                      <div style={{fontWeight:600,fontSize:13}}>{username} <span style={{fontSize:10,color:"#6b7280",fontWeight:400}}>({label})</span></div>
                      <div style={{fontSize:11,color:"#6b7280"}}>{u.nombre}{u.grado?" · "+u.grado:""}</div>
                    </div>
                    {editUser===username
                      ?<div style={{display:"flex",gap:6}}>
                          <input value={newPass} onChange={e=>setNewPass(e.target.value)} placeholder="Nueva clave"
                            style={{padding:"5px 8px",borderRadius:6,border:"1.5px solid #ddd",fontSize:12,width:120}}/>
                          <button onClick={()=>guardarPass(username)} style={{padding:"5px 10px",borderRadius:6,background:"#1565c0",color:"#fff",border:"none",cursor:"pointer",fontSize:12}}>Guardar</button>
                          <button onClick={()=>setEditUser(null)} style={{padding:"5px 10px",borderRadius:6,background:"#e9ecef",border:"none",cursor:"pointer",fontSize:12}}>✕</button>
                        </div>
                      :<button onClick={()=>{setEditUser(username);setNewPass("");}} style={{padding:"5px 12px",borderRadius:6,background:"#e9ecef",border:"none",cursor:"pointer",fontSize:12}}>🔑 Cambiar clave</button>
                    }
                  </div>
                );
              })}
            </div>
          )}
          {subTab==="estudiantes"&&<NuevoEstudianteTab datosEscuela={datosEscuela} onDataUpdate={onDataUpdate}/>}
          {subTab==="tipos"&&<TiposTab datosEscuela={datosEscuela} onDataUpdate={onDataUpdate}/>}
          {subTab==="simat"&&(
            <div>
              <div style={{background:"#e8f4fd",border:"2px dashed #1565c0",borderRadius:12,padding:24,textAlign:"center"}}>
                <div style={{fontSize:40,marginBottom:8}}>📊</div>
                <div style={{fontWeight:700,fontSize:15,color:"#1e3a5f",marginBottom:6}}>Cargar nuevo SIMAT</div>
                <div style={{fontSize:12,color:"#6b7280",marginBottom:16}}>Archivo Excel con hojas <b>SIMAT</b> y <b>Tipos</b>.</div>
                <label style={{display:"inline-block",padding:"10px 24px",background:"#1565c0",color:"#fff",borderRadius:8,cursor:"pointer",fontWeight:700,fontSize:13}}>
                  📂 Seleccionar archivo .xlsx
                  <input type="file" accept=".xlsx,.xls" onChange={handleSIMAT} style={{display:"none"}}/>
                </label>
                {simatMsg&&<div style={{marginTop:14,padding:"10px 16px",borderRadius:8,fontSize:13,fontWeight:600,
                  background:simatMsg.startsWith("✅")?"#d4edda":simatMsg.startsWith("❌")?"#f8d7da":"#fff3cd",
                  color:simatMsg.startsWith("✅")?"#1e7e34":simatMsg.startsWith("❌")?"#721c24":"#856404"}}>{simatMsg}</div>}
              </div>
            </div>
          )}
          {subTab==="borrar"&&(
            <div>
              <div style={{background:"#fff5f5",border:"2px solid #dc354530",borderRadius:12,padding:24,textAlign:"center"}}>
                <div style={{fontSize:48,marginBottom:8}}>⚠️</div>
                <div style={{fontWeight:700,fontSize:16,color:"#721c24",marginBottom:8}}>Borrar Base de Datos</div>
                <div style={{fontSize:13,color:"#6b7280",marginBottom:20,lineHeight:1.6}}>Elimina todos los estados, asistencia, notas y entregas.<br/><b>Acción irreversible.</b></div>
                {!confirmarBorrar
                  ?<button onClick={()=>setConfirmarBorrar(true)} style={{padding:"10px 28px",borderRadius:8,background:"#dc3545",color:"#fff",fontWeight:700,fontSize:14,border:"none",cursor:"pointer"}}>🗑️ Borrar todos los datos</button>
                  :<div>
                    <div style={{fontSize:13,fontWeight:600,color:"#721c24",marginBottom:12}}>¿Confirmas? Esta acción es irreversible.</div>
                    <div style={{display:"flex",gap:12,justifyContent:"center"}}>
                      <button onClick={borrarBD} style={{padding:"10px 24px",borderRadius:8,background:"#dc3545",color:"#fff",fontWeight:700,fontSize:14,border:"none",cursor:"pointer"}}>✅ Sí, borrar todo</button>
                      <button onClick={()=>setConfirmarBorrar(false)} style={{padding:"10px 24px",borderRadius:8,background:"#e9ecef",color:"#374151",fontWeight:700,fontSize:14,border:"none",cursor:"pointer"}}>❌ Cancelar</button>
                    </div>
                  </div>
                }
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ══════════ REPORTES ══════════
function ReportesTab({datosEscuela,appState}){
  const [filtroGrado,setFiltroGrado]=useState("TODOS");
  const [filtroEstado,setFiltroEstado]=useState("TODOS");
  const [filtroSimat,setFiltroSimat]=useState("TODOS");
  const [filtroRuta,setFiltroRuta]=useState("TODOS");
  const [vista,setVista]=useState("dashboard");

  const todosEst=GRADOS_ORDEN.filter(g=>datosEscuela[g]).flatMap(grado=>{
    const ests=datosEscuela[grado].estudiantes;
    const st=appState?.[grado]?.comisiones||{};
    return ests.map((est,i)=>({...est,grado,tutor:datosEscuela[grado].tutor,estadoComision:st[i]?.estado||est.estado||"MATRICULADO"}));
  });

  const estF=todosEst.filter(e=>{
    if(filtroGrado!=="TODOS"&&e.grado!==filtroGrado)return false;
    if(filtroEstado!=="TODOS"&&e.estadoComision!==filtroEstado)return false;
    if(filtroSimat!=="TODOS"&&(e.simat||"NO REGISTRADO")!==filtroSimat)return false;
    if(filtroRuta==="CON RUTA"&&!e.ruta)return false;
    if(filtroRuta==="SIN RUTA"&&e.ruta)return false;
    return true;
  });

  const total=todosEst.length;
  const porEstado={};ESTADOS.forEach(e=>{porEstado[e.key]=todosEst.filter(x=>x.estadoComision===e.key).length;});
  const porGenero={M:todosEst.filter(e=>e.genero!=="F").length,F:todosEst.filter(e=>e.genero==="F").length};
  const conRuta=todosEst.filter(e=>e.ruta).length;
  const porSimat={};SIMAT_ESTADOS.forEach(s=>{porSimat[s.key]=todosEst.filter(e=>(e.simat||"NO REGISTRADO")===s.key).length;});
  const rutasUnicas=[...new Set(todosEst.filter(e=>e.ruta).map(e=>e.ruta))].sort();

  function BarChart({data}){
    const max=Math.max(...data.map(d=>d.val),1);
    return(<div style={{display:"flex",flexDirection:"column",gap:6}}>
      {data.map((d,i)=>(
        <div key={i} style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{fontSize:11,color:"#374151",minWidth:150,textAlign:"right"}}>{d.label}</div>
          <div style={{flex:1,background:"#f0f4f8",borderRadius:4,height:24,position:"relative"}}>
            <div style={{width:(d.val/max*100)+"%",background:d.color,height:"100%",borderRadius:4,minWidth:d.val>0?8:0,transition:"width .4s"}}/>
            <span style={{position:"absolute",left:8,top:"50%",transform:"translateY(-50%)",fontSize:11,fontWeight:700,color:"#fff",textShadow:"0 1px 2px rgba(0,0,0,.5)"}}>
              {d.val} ({pct(d.val,total)})
            </span>
          </div>
        </div>
      ))}
    </div>);
  }

  const KPI=({val,label,sub,bg1,bg2})=>(
    <div style={{background:"linear-gradient(135deg,"+bg1+","+bg2+")",borderRadius:12,padding:"16px 14px",textAlign:"center",color:"#fff"}}>
      <div style={{fontSize:32,fontWeight:800}}>{val}</div>
      <div style={{fontSize:10,opacity:.85,marginTop:2}}>{label}</div>
      {sub&&<div style={{fontSize:12,fontWeight:700,marginTop:2}}>{sub}</div>}
    </div>
  );

  return(
    <div>
      <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap"}}>
        {[["dashboard","📊 Dashboard"],["tabla","📋 Tabla detallada"],["ruta","🚌 Rutas"]].map(([k,l])=>(
          <button key={k} onClick={()=>setVista(k)}
            style={{padding:"7px 16px",borderRadius:8,border:"none",cursor:"pointer",fontSize:13,fontWeight:vista===k?700:500,
              background:vista===k?"#1565c0":"#e8edf5",color:vista===k?"#fff":"#374151"}}>{l}</button>
        ))}
      </div>

      {vista==="dashboard"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(140px,1fr))",gap:12,marginBottom:20}}>
            <KPI val={total} label="TOTAL ESTUDIANTES" bg1="#1565c0" bg2="#0d47a1"/>
            <KPI val={porEstado["MATRICULADO"]||0} label="MATRICULADOS" sub={pct(porEstado["MATRICULADO"]||0,total)} bg1="#1e7e34" bg2="#155724"/>
            <KPI val={porEstado["PROMOVIDO"]||0} label="PROMOVIDOS" sub={pct(porEstado["PROMOVIDO"]||0,total)} bg1="#004085" bg2="#002752"/>
            <KPI val={porEstado["RETIRADO"]||0} label="RETIRADOS" sub={pct(porEstado["RETIRADO"]||0,total)} bg1="#721c24" bg2="#4e1217"/>
            <KPI val={porEstado["REPROBADO"]||0} label="REPROBADOS" sub={pct(porEstado["REPROBADO"]||0,total)} bg1="#856404" bg2="#533f03"/>
            <KPI val={(porEstado["COMPROMISO CONVIVENCIAL"]||0)+(porEstado["COMPROMISO ACADEMICO"]||0)} label="EN COMPROMISO" sub={pct((porEstado["COMPROMISO CONVIVENCIAL"]||0)+(porEstado["COMPROMISO ACADEMICO"]||0),total)} bg1="#5a1e6b" bg2="#3b1247"/>
            <KPI val={porEstado["CONTINUIDAD"]||0} label="CONTINUIDAD" sub={pct(porEstado["CONTINUIDAD"]||0,total)} bg1="#495057" bg2="#212529"/>
            <KPI val={conRuta} label="CON RUTA" sub={pct(conRuta,total)} bg1="#c01565" bg2="#7a0d41"/>
            <KPI val={porGenero.M} label="HOMBRES" sub={pct(porGenero.M,total)} bg1="#1565c0" bg2="#0a3f82"/>
            <KPI val={porGenero.F} label="MUJERES" sub={pct(porGenero.F,total)} bg1="#c01565" bg2="#7a0d41"/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
            <div style={{background:"#fff",borderRadius:12,padding:18,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
              <div style={{fontWeight:700,fontSize:13,color:"#1e3a5f",marginBottom:14}}>📊 Por Estado de Comisión</div>
              <BarChart data={ESTADOS.map(e=>({label:e.label,val:porEstado[e.key]||0,color:e.color}))}/>
            </div>
            <div style={{background:"#fff",borderRadius:12,padding:18,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
              <div style={{fontWeight:700,fontSize:13,color:"#1e3a5f",marginBottom:14}}>📋 Estado SIMAT</div>
              <BarChart data={SIMAT_ESTADOS.map(s=>({label:s.label,val:porSimat[s.key]||0,color:s.color}))}/>
              <div style={{fontWeight:700,fontSize:13,color:"#1e3a5f",margin:"18px 0 10px"}}>🚌 Ruta</div>
              <BarChart data={[{label:"Con Ruta",val:conRuta,color:"#c01565"},{label:"Sin Ruta",val:total-conRuta,color:"#6b7280"}]}/>
            </div>
          </div>
          <div style={{background:"#fff",borderRadius:12,padding:18,boxShadow:"0 1px 6px rgba(0,0,0,.08)",marginBottom:16}}>
            <div style={{fontWeight:700,fontSize:13,color:"#1e3a5f",marginBottom:14}}>🏫 Estudiantes por Grado</div>
            <BarChart data={GRADOS_ORDEN.filter(g=>datosEscuela[g]).map(g=>({label:g,val:datosEscuela[g].estudiantes.length,color:"#1565c0"}))}/>
          </div>
          {(porEstado["RETIRADO"]||0)>0&&(
            <div style={{background:"#fff",borderRadius:12,padding:18,boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
              <div style={{fontWeight:700,fontSize:13,color:"#721c24",marginBottom:12}}>🚨 Estudiantes Retirados ({porEstado["RETIRADO"]})</div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#f8d7da"}}>
                  <th style={th}>#</th><th style={{...th,textAlign:"left"}}>Nombre</th><th style={th}>Grado</th><th style={{...th,textAlign:"left"}}>Tutor</th><th style={th}>Gén.</th>
                </tr></thead>
                <tbody>{todosEst.filter(e=>e.estadoComision==="RETIRADO").map((e,i)=>(
                  <tr key={i} style={{borderBottom:"1px solid #f0f4f8",background:i%2===0?"#fff":"#fff5f5"}}>
                    <td style={{...td,textAlign:"center",color:"#9ca3af"}}>{i+1}</td>
                    <td style={{...td,color:"#721c24",fontWeight:500,textDecoration:"line-through"}}>{e.nombre}</td>
                    <td style={{...td,textAlign:"center"}}>{e.grado}</td>
                    <td style={td}>{e.tutor}</td>
                    <td style={{...td,textAlign:"center"}}>{e.genero}</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {vista==="tabla"&&(
        <div>
          <div style={{background:"#fff",borderRadius:12,padding:16,marginBottom:16,boxShadow:"0 1px 6px rgba(0,0,0,.08)",display:"flex",gap:12,flexWrap:"wrap",alignItems:"flex-end"}}>
            {[
              ["Grado",filtroGrado,setFiltroGrado,[["TODOS","Todos los grados"],...GRADOS_ORDEN.filter(g=>datosEscuela[g]).map(g=>[g,g])]],
              ["Estado Comisión",filtroEstado,setFiltroEstado,[["TODOS","Todos"],...ESTADOS.map(e=>[e.key,e.label])]],
              ["Estado SIMAT",filtroSimat,setFiltroSimat,[["TODOS","Todos"],...SIMAT_ESTADOS.map(s=>[s.key,s.label])]],
              ["Ruta",filtroRuta,setFiltroRuta,[["TODOS","Todos"],["CON RUTA","Con ruta"],["SIN RUTA","Sin ruta"]]],
            ].map(([label,val,setter,opts])=>(
              <div key={label}>
                <label style={{fontSize:11,fontWeight:600,color:"#374151",display:"block",marginBottom:3}}>{label}</label>
                <select value={val} onChange={e=>setter(e.target.value)} style={{padding:"6px 12px",borderRadius:7,border:"1.5px solid #ddd",fontSize:12}}>
                  {opts.map(([k,l])=><option key={k} value={k}>{l}</option>)}
                </select>
              </div>
            ))}
            <div style={{marginLeft:"auto",fontSize:13,fontWeight:600,color:"#1565c0",alignSelf:"center"}}>{estF.length} de {total}</div>
          </div>
          <div style={{background:"#fff",borderRadius:12,overflow:"auto",boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{background:"#1565c0",color:"#fff"}}>
                {["#","Apellidos y Nombres","Grado","Gén.","Estado Comisión","SIMAT","Ruta","Transportador","F. Nacimiento"].map(h=>(
                  <th key={h} style={{...th,color:"#fff",textAlign:h==="Apellidos y Nombres"?"left":"center"}}>{h}</th>
                ))}
              </tr></thead>
              <tbody>{estF.map((e,i)=>{
                const info=estadoInfo(e.estadoComision),si=simatInfo(e.simat||"NO REGISTRADO");
                return(<tr key={i} style={{background:i%2===0?"#fff":"#f9fbff",borderBottom:"1px solid #e9ecf0"}}>
                  <td style={{...td,textAlign:"center",color:"#9ca3af"}}>{i+1}</td>
                  <td style={{...td,fontWeight:500}}>{e.nombre}</td>
                  <td style={{...td,textAlign:"center",fontSize:11}}>{e.grado}</td>
                  <td style={{...td,textAlign:"center"}}>{e.genero}</td>
                  <td style={td}><span style={{background:info.bg,color:info.color,padding:"2px 8px",borderRadius:5,fontSize:11,fontWeight:600}}>{info.label}</span></td>
                  <td style={td}><span style={{background:si.bg,color:si.color,padding:"2px 8px",borderRadius:5,fontSize:11,fontWeight:600}}>{si.label}</span></td>
                  <td style={{...td,fontSize:11}}>{e.ruta||<span style={{color:"#d1d5db"}}>–</span>}</td>
                  <td style={{...td,fontSize:11}}>{e.transportador||<span style={{color:"#d1d5db"}}>–</span>}</td>
                  <td style={{...td,textAlign:"center",fontSize:11}}>{e.fechaNac||<span style={{color:"#d1d5db"}}>–</span>}</td>
                </tr>);
              })}</tbody>
            </table>
          </div>
        </div>
      )}

      {vista==="ruta"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(150px,1fr))",gap:12,marginBottom:20}}>
            <KPI val={conRuta} label="CON RUTA" sub={pct(conRuta,total)} bg1="#c01565" bg2="#7a0d41"/>
            <KPI val={total-conRuta} label="SIN RUTA" sub={pct(total-conRuta,total)} bg1="#495057" bg2="#212529"/>
            <KPI val={rutasUnicas.length} label="RUTAS DISTINTAS" bg1="#1565c0" bg2="#0d47a1"/>
          </div>
          {rutasUnicas.length===0&&<div style={{background:"#fff",borderRadius:12,padding:32,textAlign:"center",color:"#9ca3af",fontSize:13}}>No hay estudiantes con ruta asignada.</div>}
          {rutasUnicas.map(ruta=>{
            const ests=todosEst.filter(e=>e.ruta===ruta);
            const trans=[...new Set(ests.map(e=>e.transportador).filter(Boolean))];
            return(<div key={ruta} style={{background:"#fff",borderRadius:12,marginBottom:12,overflow:"hidden",boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
              <div style={{background:"#c01565",color:"#fff",padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontWeight:700,fontSize:13}}>🚌 {ruta}</span>
                <span style={{fontSize:12}}>{ests.length} estudiantes · {trans.join(", ")||"Sin transportador"}</span>
              </div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr style={{background:"#f9f0f5"}}>
                  <th style={th}>#</th><th style={{...th,textAlign:"left"}}>Nombre</th><th style={th}>Grado</th><th style={th}>Gén.</th><th style={th}>Transportador</th>
                </tr></thead>
                <tbody>{ests.map((e,i)=>(
                  <tr key={i} style={{borderBottom:"1px solid #f0f4f8",background:i%2===0?"#fff":"#fdf8fb"}}>
                    <td style={{...td,textAlign:"center",color:"#9ca3af"}}>{i+1}</td>
                    <td style={{...td,fontWeight:500}}>{e.nombre}</td>
                    <td style={{...td,textAlign:"center",fontSize:11}}>{e.grado}</td>
                    <td style={{...td,textAlign:"center"}}>{e.genero}</td>
                    <td style={{...td,fontSize:11}}>{e.transportador||"–"}</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>);
          })}
        </div>
      )}
    </div>
  );
}

// ══════════ APP PRINCIPAL ══════════
export default function App(){
  const [auth,setAuth]=useState(()=>loadStorage(AUTH_KEY,null));
  const [datosEscuela,setDatosEscuela]=useState(()=>loadStorage(DATA_KEY,DATOS_ESCUELA_INICIAL));
  const [tab,setTab]=useState("comisiones");
  const [gradoSel,setGradoSel]=useState(()=>{const a=loadStorage(AUTH_KEY,null);return a?.grado||GRADOS_ORDEN.find(g=>loadStorage(DATA_KEY,DATOS_ESCUELA_INICIAL)[g])||"";});
  const [appState,setAppState]=useState(()=>loadStorage(STORAGE_KEY,{}));
  const [showAdmin,setShowAdmin]=useState(false);

  useEffect(()=>{saveStorage(STORAGE_KEY,appState);},[appState]);

  function handleLogin(session){setAuth(session);if(session.grado)setGradoSel(session.grado);else setGradoSel(GRADOS_ORDEN.find(g=>datosEscuela[g])||"");}
  function handleLogout(){saveStorage(AUTH_KEY,null);setAuth(null);}
  const update=useCallback((fn)=>{setAppState(prev=>fn(JSON.parse(JSON.stringify(prev))));},[]);

  if(!auth)return <LoginScreen onLogin={handleLogin}/>;

  const isAdmin=auth.role==="admin",isCoord=auth.role==="coordinador";
  const grados=GRADOS_ORDEN.filter(g=>datosEscuela[g]);
  const gradoData=datosEscuela[gradoSel]||{tutor:"",estudiantes:[]};
  const estudiantes=gradoData.estudiantes;

  function getEstado(idx){return appState?.[gradoSel]?.comisiones?.[idx]?.estado||estudiantes[idx]?.estado||"MATRICULADO";}
  function setEstado(idx,val){update(p=>{if(!p[gradoSel])p[gradoSel]={};if(!p[gradoSel].comisiones)p[gradoSel].comisiones={};if(!p[gradoSel].comisiones[idx])p[gradoSel].comisiones[idx]={};p[gradoSel].comisiones[idx].estado=val;return p;});}
  function getAsistencia(idx,s,d){return appState?.[gradoSel]?.asistencia?.[idx]?.[s]?.[d]||"";}
  function toggleAsistencia(idx,s,d){update(p=>{if(!p[gradoSel])p[gradoSel]={};if(!p[gradoSel].asistencia)p[gradoSel].asistencia={};if(!p[gradoSel].asistencia[idx])p[gradoSel].asistencia[idx]={};if(!p[gradoSel].asistencia[idx][s])p[gradoSel].asistencia[idx][s]={};const c=p[gradoSel].asistencia[idx][s][d]||"";p[gradoSel].asistencia[idx][s][d]=c==="P"?"F":c==="F"?"":"P";return p;});}
  function getNota(idx,per,sab){return appState?.[gradoSel]?.notas?.[idx]?.[per]?.[sab]??"";}
  function setNota(idx,per,sab,v){update(p=>{if(!p[gradoSel])p[gradoSel]={};if(!p[gradoSel].notas)p[gradoSel].notas={};if(!p[gradoSel].notas[idx])p[gradoSel].notas[idx]={};if(!p[gradoSel].notas[idx][per])p[gradoSel].notas[idx][per]={};p[gradoSel].notas[idx][per][sab]=v;return p;});}
  function getEntrega(idx){return appState?.[gradoSel]?.entregas?.[idx]??false;}
  function toggleEntrega(idx){update(p=>{if(!p[gradoSel])p[gradoSel]={};if(!p[gradoSel].entregas)p[gradoSel].entregas={};p[gradoSel].entregas[idx]=!p[gradoSel].entregas[idx];return p;});}
  function getSimat(idx){return appState?.[gradoSel]?.simat?.[idx]??estudiantes[idx]?.simat??"NO REGISTRADO";}
  function setSimat(idx,val){update(p=>{if(!p[gradoSel])p[gradoSel]={};if(!p[gradoSel].simat)p[gradoSel].simat={};p[gradoSel].simat[idx]=val;return p;});}
  function calcProm(idx,per){const v=SABERES.map(s=>parseFloat(getNota(idx,per,s))).filter(n=>!isNaN(n));return v.length?(v.reduce((a,b)=>a+b,0)/v.length).toFixed(1):"";}
  function calcPromGeneral(idx){const v=PERIODOS.flatMap(p=>SABERES.map(s=>parseFloat(getNota(idx,p,s)))).filter(n=>!isNaN(n));return v.length?(v.reduce((a,b)=>a+b,0)/v.length).toFixed(1):"";}
  function contarAsistencia(idx){let p=0,f=0;SEMANAS.forEach(s=>DIAS.forEach((_,di)=>{const v=getAsistencia(idx,s,di.toString());if(v==="P")p++;else if(v==="F")f++;}));return{p,f};}

  const contadores={};ESTADOS.forEach(e=>{contadores[e.key]=0;});
  estudiantes.forEach((_,i)=>{const e=getEstado(i);if(contadores[e]!==undefined)contadores[e]++;});

  function handlePrint(){
    const titulos={comisiones:"Acta de Comisión Final",asistencia:"Asistencia y Notas",entregas:"Asistencia Entrega de Informes",consolidado:"Consolidado Final"};
    const ps=`<style>@page{size:landscape;margin:1cm}body{font-family:'Segoe UI',sans-serif;font-size:10px}.ph{text-align:center;margin-bottom:12px}.ph h1{font-size:13px;margin:0 0 2px}.ph h2{font-size:11px;margin:0 0 2px;font-weight:600}.ph p{font-size:10px;margin:0;color:#444}table{width:100%;border-collapse:collapse;font-size:9px}th{background:#1565c0!important;color:#fff!important;padding:5px 6px;border:1px solid #ccc;text-align:center}td{padding:4px 6px;border:1px solid #ccc;vertical-align:middle}tr:nth-child(even) td{background:#f5f8ff}.ret td{background:#fff0f0!important;text-decoration:line-through;color:#c00}.badge{display:inline-block;padding:2px 6px;border-radius:4px;font-size:8px;font-weight:700}.counters{display:flex;gap:8px;margin-bottom:10px;flex-wrap:wrap}.counter{border:1px solid #ccc;border-radius:6px;padding:4px 10px;text-align:center;min-width:80px}.counter .num{font-size:18px;font-weight:700}.counter .lbl{font-size:8px;font-weight:600}.pk{background:#d4edda!important;color:#1e7e34;font-weight:700}.fk{background:#f8d7da!important;color:#721c24;font-weight:700}.pok{color:#1e7e34;font-weight:700}.pmal{color:#c00;font-weight:700}</style>`;
    let body="";
    if(tab==="comisiones"){
      const cnt=ESTADOS.map(e=>`<div class="counter"><div class="num" style="color:${e.color}">${contadores[e.key]}</div><div class="lbl" style="color:${e.color}">${e.label.toUpperCase()}</div></div>`).join("");
      const hc=ESTADOS.slice(1).map(e=>`<th style="font-size:8px">${e.label}</th>`).join("");
      const rows=estudiantes.map((est,i)=>{const es=getEstado(i),info=estadoInfo(es),ret=es==="RETIRADO";return`<tr class="${ret?"ret":""}"><td style="text-align:center">${i+1}</td><td>${est.nombre}</td><td style="text-align:center">${est.genero}</td><td style="text-align:center"><span class="badge" style="background:${info.bg};color:${info.color}">${info.label}</span></td>${ESTADOS.slice(1).map(e=>`<td style="text-align:center">${es===e.key?"✓":"–"}</td>`).join("")}</tr>`;}).join("");
      body=`<div class="counters">${cnt}</div><table><thead><tr><th>#</th><th style="text-align:left">Apellidos y Nombres</th><th>Gén.</th><th>Estado</th>${hc}</tr></thead><tbody>${rows}</tbody></table>`;
    }
    if(tab==="asistencia"){
      const sc=SEMANAS.flatMap(s=>DIAS_LABELS.map(d=>`<th>${s}-${d[0]}</th>`)).join("");
      const nc=PERIODOS.flatMap(p=>SABERES.map(s=>`<th>${p} ${s.slice(0,5)}</th>`)).join("");
      const rows=estudiantes.map((est,i)=>{const{p,f}=contarAsistencia(i);const ac=SEMANAS.flatMap(s=>DIAS.map((_,di)=>{const v=getAsistencia(i,s,di.toString());return`<td class="${v==="P"?"pk":v==="F"?"fk":""}" style="text-align:center">${v||""}</td>`;})).join("");const nc2=PERIODOS.flatMap(per=>SABERES.map(sab=>`<td style="text-align:center">${getNota(i,per,sab)}</td>`)).join("");const pm=PERIODOS.map(per=>{const pv=calcProm(i,per);return`<td class="${pv&&parseFloat(pv)>=3?"pok":"pmal"}" style="text-align:center">${pv}</td>`;}).join("");const pg=calcPromGeneral(i);return`<tr><td style="text-align:center">${i+1}</td><td>${est.nombre}</td>${ac}<td style="text-align:center;font-weight:700;color:#1e7e34">${p}</td><td style="text-align:center;font-weight:700;color:#c00">${f}</td>${nc2}${pm}<td class="${pg&&parseFloat(pg)>=3?"pok":"pmal"}" style="text-align:center">${pg}</td></tr>`;}).join("");
      body=`<table><thead><tr><th>#</th><th style="text-align:left;min-width:130px">Apellidos y Nombres</th>${sc}<th>P</th><th>F</th>${nc}${PERIODOS.map(p=>`<th>${p} PM</th>`).join("")}<th>PM Gral</th></tr></thead><tbody>${rows}</tbody></table>`;
    }
    if(tab==="entregas"){
      const asi=estudiantes.filter((_,i)=>getEntrega(i)).length;
      const rows=estudiantes.map((est,i)=>{const a=getEntrega(i);return`<tr><td style="text-align:center">${i+1}</td><td>${est.nombre}</td><td style="text-align:center">${est.genero}</td><td style="text-align:center;font-weight:700;color:${a?"#1e7e34":"#c00"}">${a?"✓ Sí":"✗ No"}</td><td></td><td>${est.telefono||""}</td></tr>`;}).join("");
      body=`<div class="counters"><div class="counter"><div class="num" style="color:#1e7e34">${asi}</div><div class="lbl" style="color:#1e7e34">ASISTIERON</div></div><div class="counter"><div class="num" style="color:#c00">${estudiantes.length-asi}</div><div class="lbl" style="color:#c00">NO ASIST.</div></div></div><table><thead><tr><th>#</th><th style="text-align:left">Apellidos y Nombres</th><th>Gén.</th><th>Asistió</th><th style="min-width:120px">Firma</th><th>Teléfono</th></tr></thead><tbody>${rows}</tbody></table>`;
    }
    if(tab==="consolidado"){
      const filas=GRADOS_ORDEN.filter(g=>datosEscuela[g]).map(grado=>{const ests=datosEscuela[grado].estudiantes,st=appState?.[grado]?.comisiones||{},counts={};ESTADOS.forEach(e=>{counts[e.key]={M:0,F:0};});ests.forEach((est,i)=>{const estado=st[i]?.estado||est.estado||"MATRICULADO",gen=est.genero==="F"?"F":"M";if(counts[estado])counts[estado][gen]++;});return{grado,tutor:datosEscuela[grado].tutor,counts,totalM:ests.filter(e=>e.genero!=="F").length,totalF:ests.filter(e=>e.genero==="F").length,total:ests.length};});
      const tot={M:0,F:0,total:0},totE={};ESTADOS.forEach(e=>{totE[e.key]={M:0,F:0};});filas.forEach(f=>{tot.M+=f.totalM;tot.F+=f.totalF;tot.total+=f.total;ESTADOS.forEach(e=>{totE[e.key].M+=f.counts[e.key].M;totE[e.key].F+=f.counts[e.key].F;});});
      const hc=ESTADOS.map(e=>`<th colspan="2" style="background:${e.color};color:#fff;font-size:8px">${e.label}</th>`).join("");
      const rows=filas.map((f,ri)=>{const cols=ESTADOS.map(e=>`<td style="text-align:center;${f.counts[e.key].M>0?`background:${e.bg};color:${e.color};font-weight:700`:""}">${f.counts[e.key].M||"–"}</td><td style="text-align:center;${f.counts[e.key].F>0?`background:${e.bg};color:${e.color};font-weight:700`:""}">${f.counts[e.key].F||"–"}</td>`).join("");return`<tr style="${ri%2===0?"":"background:#f5f8ff"}"><td style="font-weight:600">${f.grado}</td><td style="font-size:9px">${f.tutor}</td><td style="text-align:center;color:#1565c0;font-weight:700">${f.totalM}</td><td style="text-align:center;color:#c01565;font-weight:700">${f.totalF}</td><td style="text-align:center;font-weight:700">${f.total}</td>${cols}</tr>`;}).join("");
      const tr=ESTADOS.map(e=>`<td style="text-align:center;color:#7dd3fc;font-weight:700">${totE[e.key].M||"–"}</td><td style="text-align:center;color:#f9a8d4;font-weight:700">${totE[e.key].F||"–"}</td>`).join("");
      body=`<table><thead><tr style="background:#1e3a5f;color:#fff"><th colspan="5" style="text-align:left;color:#fff">Grado / Tutor</th>${hc}</tr><tr style="background:#e8edf5"><th style="text-align:left">Grado</th><th style="text-align:left">Tutor(a)</th><th style="color:#1565c0">♂ H</th><th style="color:#c01565">♀ M</th><th>Total</th>${ESTADOS.map(()=>"<th>♂</th><th>♀</th>").join("")}</tr></thead><tbody>${rows}<tr style="background:#1e3a5f;color:#fff;font-weight:700"><td colspan="2" style="color:#fff">TOTALES</td><td style="text-align:center;color:#7dd3fc">${tot.M}</td><td style="text-align:center;color:#f9a8d4">${tot.F}</td><td style="text-align:center;color:#fff">${tot.total}</td>${tr}</tr></tbody></table>`;
    }
    const titulo=titulos[tab]||"Reporte";
    const win=window.open("","_blank");
    win.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"/>${ps}</head><body><div class="ph"><h1>IED SARID ARTETA DE VÁSQUEZ — INEDISAV</h1><h2>${titulo}</h2><p>${tab==="consolidado"?"Año Escolar 2026 · Todos los Grados":`Grado: <strong>${gradoSel}</strong> · Director(a): <strong>${gradoData.tutor}</strong> · 2026`}</p></div>${body}</body></html>`);
    win.document.close();win.focus();setTimeout(()=>win.print(),400);
  }

  const TABS_MAIN=[
    {key:"comisiones",label:"📋 Comisiones"},
    {key:"asistencia",label:"📅 Asistencia y Notas",adminOnly:true},
    {key:"entregas",label:"📦 Entrega de Informes",adminOnly:true},
    {key:"consolidado",label:"📊 Consolidado Final"},
    {key:"reportes",label:"📈 Reportes",adminOnly:true,coordOk:true},
  ];

  return(
    <div style={{fontFamily:"'Segoe UI',sans-serif",minHeight:"100vh",background:"#f0f4f8"}}>
      {showAdmin&&<AdminPanel onClose={()=>setShowAdmin(false)} datosEscuela={datosEscuela} onDataUpdate={d=>setDatosEscuela(d)}/>}
      <div style={{background:"linear-gradient(135deg,#0d47a1,#1565c0)",color:"#fff",padding:"12px 24px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:10,opacity:.75,letterSpacing:2,textTransform:"uppercase"}}>INEDISAV · 2026</div>
          <div style={{fontSize:18,fontWeight:700}}>IED Sarid Arteta de Vásquez</div>
          <div style={{fontSize:11,opacity:.8}}>Sistema de Gestión Académica</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:12,fontWeight:600}}>{auth.nombre}</div>
            <div style={{fontSize:10,opacity:.75}}>{isAdmin?"Administrador":isCoord?"Coordinación":auth.grado}</div>
          </div>
          {isAdmin&&<button onClick={()=>setShowAdmin(true)} style={{padding:"6px 12px",borderRadius:8,background:"rgba(255,255,255,.2)",border:"1px solid rgba(255,255,255,.4)",color:"#fff",cursor:"pointer",fontSize:12,fontWeight:600}}>⚙️ Admin</button>}
          <button onClick={handleLogout} style={{padding:"6px 12px",borderRadius:8,background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.3)",color:"#fff",cursor:"pointer",fontSize:12}}>🚪 Salir</button>
        </div>
      </div>

      {tab!=="reportes"&&(
        <div style={{background:"#fff",borderBottom:"2px solid #e3e8f0",padding:"10px 24px",display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
          <label style={{fontWeight:600,fontSize:13,color:"#374151"}}>📚 Grado:</label>
          {(isAdmin||isCoord)
            ?<select value={gradoSel} onChange={e=>setGradoSel(e.target.value)} style={{padding:"6px 14px",borderRadius:8,border:"1.5px solid #bdc3cf",fontSize:14,fontWeight:600,color:"#1e3a5f",background:"#f8faff",cursor:"pointer"}}>
                {grados.map(g=><option key={g} value={g}>{g}</option>)}
              </select>
            :<span style={{padding:"6px 14px",borderRadius:8,background:"#f0f4f8",fontWeight:700,color:"#1e3a5f",fontSize:14}}>{gradoSel}</span>
          }
          <div style={{fontSize:13,color:"#6b7280"}}>👤 <strong>{gradoData.tutor||"Sin tutor"}</strong> · {estudiantes.length} estudiantes</div>
        </div>
      )}

      <div style={{background:"#fff",borderBottom:"2px solid #e3e8f0",padding:"0 24px",display:"flex",gap:4,alignItems:"center",overflowX:"auto"}}>
        {TABS_MAIN.filter(t=>!t.adminOnly||isAdmin||(isCoord&&t.coordOk)).map(t=>(
          <button key={t.key} onClick={()=>setTab(t.key)}
            style={{padding:"12px 16px",border:"none",background:"none",cursor:"pointer",fontWeight:tab===t.key?700:400,fontSize:13,whiteSpace:"nowrap",
              color:tab===t.key?"#1565c0":"#6b7280",borderBottom:tab===t.key?"3px solid #1565c0":"3px solid transparent",marginBottom:-2,transition:"all .15s"}}>
            {t.label}
          </button>
        ))}
        {tab!=="reportes"&&(
          <div style={{marginLeft:"auto"}}>
            <button onClick={handlePrint} style={{padding:"7px 18px",borderRadius:8,border:"1.5px solid #1565c0",background:"#1565c0",color:"#fff",fontWeight:700,fontSize:13,cursor:"pointer"}}>🖨️ Imprimir</button>
          </div>
        )}
      </div>

      <div style={{padding:"20px 24px",maxWidth:1400,margin:"0 auto"}}>

        {tab==="comisiones"&&(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:10,marginBottom:20}}>
              {ESTADOS.map(e=>(
                <div key={e.key} style={{background:e.bg,border:`2px solid ${e.color}20`,borderRadius:10,padding:"12px 14px",textAlign:"center"}}>
                  <div style={{fontSize:26,fontWeight:700,color:e.color}}>{contadores[e.key]}</div>
                  <div style={{fontSize:10,fontWeight:600,color:e.color}}>{e.label.toUpperCase()}</div>
                </div>
              ))}
            </div>
            <div style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
              <div style={{background:"#1565c0",color:"#fff",padding:"10px 16px",fontSize:13,fontWeight:600}}>ACTA DE COMISIÓN FINAL · {gradoSel} · Tutor: {gradoData.tutor}</div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:"#e8edf5"}}>
                  <th style={th}>#</th>
                  <th style={{...th,textAlign:"left"}}>Apellidos y Nombres</th>
                  <th style={th}>Gén.</th>
                  <th style={{...th,minWidth:180}}>Estado Comisión</th>
                  {ESTADOS.slice(1).map(e=><th key={e.key} style={{...th,fontSize:10,maxWidth:60,lineHeight:1.2}}>{e.label}</th>)}
                  {isAdmin&&<th style={{...th,minWidth:140,fontSize:10,background:"#1e3a5f",color:"#fff"}}>Estado SIMAT</th>}
                </tr></thead>
                <tbody>
                  {estudiantes.map((est,i)=>{
                    const estado=getEstado(i),info=estadoInfo(estado),ret=estado==="RETIRADO";
                    return(
                      <tr key={i} style={{background:ret?"#fff5f5":i%2===0?"#fff":"#f9fbff",borderBottom:"1px solid #e9ecf0"}}>
                        <td style={{...td,textAlign:"center",color:"#9ca3af"}}>{i+1}</td>
                        <td style={{...td,fontWeight:ret?400:500,color:ret?"#dc3545":"#1e293b",textDecoration:ret?"line-through":"none"}}>{est.nombre}</td>
                        <td style={{...td,textAlign:"center",color:"#6b7280"}}>{est.genero}</td>
                        <td style={td}>
                          <select value={estado} onChange={e=>(isAdmin||(!isCoord&&estado!=="RETIRADO"))&&setEstado(i,e.target.value)}
                            disabled={isCoord||(!isAdmin&&estado==="RETIRADO")}
                            style={{width:"100%",padding:"5px 8px",borderRadius:6,fontSize:12,fontWeight:600,border:`1.5px solid ${info.color}40`,background:info.bg,color:info.color,cursor:isCoord?"default":"pointer"}}>
                            {ESTADOS.filter(e=>isAdmin||e.key!=="RETIRADO"||estado==="RETIRADO").map(e=><option key={e.key} value={e.key}>{e.label}</option>)}
                          </select>
                        </td>
                        {ESTADOS.slice(1).map(e=>(
                          <td key={e.key} style={{...td,textAlign:"center"}}>
                            {estado===e.key?<span style={{fontSize:16}}>✓</span>:<span style={{color:"#d1d5db"}}>–</span>}
                          </td>
                        ))}
                        {isAdmin&&(()=>{const sv=getSimat(i),si=simatInfo(sv);return(
                          <td style={td}>
                            <select value={sv} onChange={e=>setSimat(i,e.target.value)}
                              style={{width:"100%",padding:"5px 8px",borderRadius:6,fontSize:11,fontWeight:600,border:`1.5px solid ${si.color}40`,background:si.bg,color:si.color,cursor:"pointer"}}>
                              {SIMAT_ESTADOS.map(s=><option key={s.key} value={s.key}>{s.label}</option>)}
                            </select>
                          </td>
                        );})()}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab==="asistencia"&&(isAdmin||isCoord)&&(
          <div style={{overflowX:"auto"}}>
            <div style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 1px 6px rgba(0,0,0,.08)",minWidth:900}}>
              <div style={{background:"#1565c0",color:"#fff",padding:"10px 16px",fontSize:13,fontWeight:600}}>ASISTENCIA Y PLANILLA DE NOTAS · {gradoSel} · Director(a): {gradoData.tutor}</div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                <thead><tr style={{background:"#e8edf5"}}>
                  <th style={{...th,fontSize:10}}>#</th>
                  <th style={{...th,textAlign:"left",minWidth:160,fontSize:10}}>Apellidos y Nombres</th>
                  {SEMANAS.map(s=>DIAS.map((_,di)=><th key={s+di} style={{...th,fontSize:9,maxWidth:28,padding:"4px 2px"}}>{s}-{DIAS_LABELS[di][0]}</th>))}
                  <th style={{...th,fontSize:10}}>P</th><th style={{...th,fontSize:10}}>F</th>
                  {PERIODOS.map(p=>SABERES.map(s=><th key={p+s} style={{...th,fontSize:9,maxWidth:36,lineHeight:1.2}}>{p}<br/>{s.slice(0,5)}</th>))}
                  {PERIODOS.map(p=><th key={"pm"+p} style={{...th,fontSize:9}}>{p}<br/>PM</th>)}
                  <th style={{...th,fontSize:9}}>PM<br/>Gral</th>
                </tr></thead>
                <tbody>
                  {estudiantes.map((est,i)=>{
                    const{p,f}=contarAsistencia(i);
                    return(
                      <tr key={i} style={{background:i%2===0?"#fff":"#f9fbff",borderBottom:"1px solid #e9ecf0"}}>
                        <td style={{...td,textAlign:"center",color:"#9ca3af"}}>{i+1}</td>
                        <td style={{...td,fontWeight:500}}>{est.nombre}</td>
                        {SEMANAS.map(s=>DIAS.map((_,di)=>{const v=getAsistencia(i,s,di.toString());return(
                          <td key={s+di} style={{...td,textAlign:"center",padding:"2px 1px"}}>
                            <button onClick={()=>isAdmin&&toggleAsistencia(i,s,di.toString())} disabled={!isAdmin}
                              style={{width:24,height:22,border:"1px solid #ddd",borderRadius:4,fontSize:10,fontWeight:700,cursor:isAdmin?"pointer":"default",
                                background:v==="P"?"#d4edda":v==="F"?"#f8d7da":"#fff",color:v==="P"?"#1e7e34":v==="F"?"#721c24":"#9ca3af"}}>
                              {v||"·"}
                            </button>
                          </td>
                        );})}
                        <td style={{...td,textAlign:"center",fontWeight:700,color:"#1e7e34"}}>{p}</td>
                        <td style={{...td,textAlign:"center",fontWeight:700,color:"#dc3545"}}>{f}</td>
                        {PERIODOS.map(per=>SABERES.map((sab,si)=>(
                          <td key={per+si} style={{...td,padding:"2px 1px"}}>
                            <input type="number" min="1" max="5" step="0.1" value={getNota(i,per,sab)} onChange={e=>setNota(i,per,sab,e.target.value)} disabled={!isAdmin}
                              style={{width:36,padding:"2px 3px",border:"1px solid #ddd",borderRadius:4,fontSize:11,textAlign:"center",background:isAdmin?"#fff":"#f9fbff"}}/>
                          </td>
                        )))}
                        {PERIODOS.map(per=>{const pm=calcProm(i,per);return(
                          <td key={"pm"+per} style={{...td,textAlign:"center",fontWeight:600,color:parseFloat(pm)>=3?"#1e7e34":"#dc3545",background:"#f0f4f8"}}>{pm}</td>
                        );})}
                        <td style={{...td,textAlign:"center",fontWeight:700,background:"#e8edf5",color:parseFloat(calcPromGeneral(i))>=3?"#1e7e34":"#dc3545"}}>{calcPromGeneral(i)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab==="entregas"&&(isAdmin||isCoord)&&(
          <div>
            <div style={{display:"flex",gap:12,marginBottom:16}}>
              {[{label:"Asistieron",val:estudiantes.filter((_,i)=>getEntrega(i)).length,color:"#1e7e34",bg:"#d4edda"},{label:"No asistieron",val:estudiantes.filter((_,i)=>!getEntrega(i)).length,color:"#721c24",bg:"#f8d7da"},{label:"Total",val:estudiantes.length,color:"#004085",bg:"#cce5ff"}].map(c=>(
                <div key={c.label} style={{background:c.bg,border:`2px solid ${c.color}30`,borderRadius:10,padding:"10px 18px",textAlign:"center",minWidth:100}}>
                  <div style={{fontSize:24,fontWeight:700,color:c.color}}>{c.val}</div>
                  <div style={{fontSize:11,color:c.color,fontWeight:600}}>{c.label}</div>
                </div>
              ))}
            </div>
            <div style={{background:"#fff",borderRadius:12,overflow:"hidden",boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
              <div style={{background:"#1565c0",color:"#fff",padding:"10px 16px",fontSize:13,fontWeight:600}}>ASISTENCIA REUNIÓN PADRES · {gradoSel} · Director(a): {gradoData.tutor}</div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:"#e8edf5"}}>
                  <th style={th}>#</th><th style={{...th,textAlign:"left"}}>Apellidos y Nombres</th><th style={th}>Género</th><th style={th}>Asistió</th><th style={{...th,textAlign:"left"}}>Teléfono</th>
                </tr></thead>
                <tbody>
                  {estudiantes.map((est,i)=>{const a=getEntrega(i);return(
                    <tr key={i} style={{background:a?"#f0fff4":i%2===0?"#fff":"#f9fbff",borderBottom:"1px solid #e9ecf0"}}>
                      <td style={{...td,textAlign:"center",color:"#9ca3af"}}>{i+1}</td>
                      <td style={{...td,fontWeight:500}}>{est.nombre}</td>
                      <td style={{...td,textAlign:"center",color:"#6b7280"}}>{est.genero}</td>
                      <td style={{...td,textAlign:"center"}}>
                        <button onClick={()=>isAdmin&&toggleEntrega(i)} disabled={!isAdmin}
                          style={{padding:"5px 16px",borderRadius:20,border:"none",cursor:isAdmin?"pointer":"default",fontWeight:700,fontSize:12,background:a?"#d4edda":"#f8d7da",color:a?"#1e7e34":"#721c24"}}>
                          {a?"✓ Sí":"✗ No"}
                        </button>
                      </td>
                      <td style={td}>{est.telefono?<a href={`tel:${est.telefono}`} style={{color:"#1565c0",textDecoration:"none"}}>📞 {est.telefono}</a>:<span style={{color:"#d1d5db"}}>–</span>}</td>
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {tab==="consolidado"&&(()=>{
          const filas=GRADOS_ORDEN.filter(g=>datosEscuela[g]).map(grado=>{
            const ests=datosEscuela[grado].estudiantes,st=appState?.[grado]?.comisiones||{},counts={};
            ESTADOS.forEach(e=>{counts[e.key]={M:0,F:0};});
            ests.forEach((est,i)=>{const estado=st[i]?.estado||est.estado||"MATRICULADO",gen=est.genero==="F"?"F":"M";if(counts[estado])counts[estado][gen]++;});
            return{grado,tutor:datosEscuela[grado].tutor,counts,totalM:ests.filter(e=>e.genero!=="F").length,totalF:ests.filter(e=>e.genero==="F").length,total:ests.length};
          });
          const totales={M:0,F:0,total:0},totE={};
          ESTADOS.forEach(e=>{totE[e.key]={M:0,F:0};});
          filas.forEach(f=>{totales.M+=f.totalM;totales.F+=f.totalF;totales.total+=f.total;ESTADOS.forEach(e=>{totE[e.key].M+=f.counts[e.key].M;totE[e.key].F+=f.counts[e.key].F;});});
          const colW={minWidth:44,textAlign:"center"};
          return(
            <div>
              <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap"}}>
                {[{val:totales.total,label:"TOTAL ESTUDIANTES",color:"#1565c0",bg:"#e8f4fd"},{val:totales.M,label:"HOMBRES",color:"#1565c0",bg:"#e8f4fd"},{val:totales.F,label:"MUJERES",color:"#c01565",bg:"#fce8f4"}].map(c=>(
                  <div key={c.label} style={{background:c.bg,border:`2px solid ${c.color}20`,borderRadius:10,padding:"12px 20px",textAlign:"center"}}>
                    <div style={{fontSize:30,fontWeight:700,color:c.color}}>{c.val}</div>
                    <div style={{fontSize:11,fontWeight:600,color:c.color}}>{c.label}</div>
                  </div>
                ))}
                {ESTADOS.slice(1).map(e=>(
                  <div key={e.key} style={{background:e.bg,border:`2px solid ${e.color}20`,borderRadius:10,padding:"12px 16px",textAlign:"center"}}>
                    <div style={{fontSize:22,fontWeight:700,color:e.color}}>{totE[e.key].M+totE[e.key].F}</div>
                    <div style={{fontSize:10,fontWeight:600,color:e.color}}>{e.label.toUpperCase()}</div>
                    <div style={{fontSize:10,color:e.color,opacity:.7}}>♂{totE[e.key].M} · ♀{totE[e.key].F}</div>
                  </div>
                ))}
              </div>
              <div style={{background:"#fff",borderRadius:12,overflow:"auto",boxShadow:"0 1px 6px rgba(0,0,0,.08)"}}>
                <div style={{background:"#1565c0",color:"#fff",padding:"10px 16px",fontSize:13,fontWeight:600}}>CONSOLIDADO FINAL · AÑO ESCOLAR 2026</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead>
                    <tr style={{background:"#1e3a5f",color:"#fff"}}>
                      <th style={{...th,textAlign:"left",minWidth:120,color:"#fff"}}>Grado</th>
                      <th style={{...th,minWidth:160,textAlign:"left",color:"#fff"}}>Tutor(a)</th>
                      <th style={{...th,...colW,color:"#fff"}}>♂ H</th>
                      <th style={{...th,...colW,color:"#fff"}}>♀ M</th>
                      <th style={{...th,...colW,color:"#fff"}}>Total</th>
                      {ESTADOS.map(e=><th key={e.key} colSpan={2} style={{...th,fontSize:10,background:e.color,color:"#fff",minWidth:80}}>{e.label}</th>)}
                    </tr>
                    <tr style={{background:"#e8edf5"}}>
                      <th style={th} colSpan={5}></th>
                      {ESTADOS.map(e=>[<th key={e.key+"m"} style={{...th,fontSize:10,color:e.color}}>♂</th>,<th key={e.key+"f"} style={{...th,fontSize:10,color:e.color}}>♀</th>])}
                    </tr>
                  </thead>
                  <tbody>
                    {filas.map((f,ri)=>(
                      <tr key={f.grado} style={{background:ri%2===0?"#fff":"#f9fbff",borderBottom:"1px solid #e9ecf0"}}>
                        <td style={{...td,fontWeight:600,color:"#1e3a5f"}}>{f.grado}</td>
                        <td style={{...td,fontSize:11,color:"#374151"}}>{f.tutor}</td>
                        <td style={{...td,textAlign:"center",fontWeight:600,color:"#1565c0"}}>{f.totalM}</td>
                        <td style={{...td,textAlign:"center",fontWeight:600,color:"#c01565"}}>{f.totalF}</td>
                        <td style={{...td,textAlign:"center",fontWeight:700}}>{f.total}</td>
                        {ESTADOS.map(e=>[
                          <td key={e.key+"m"} style={{...td,textAlign:"center",background:f.counts[e.key].M>0?e.bg:"",color:e.color,fontWeight:f.counts[e.key].M>0?700:400}}>{f.counts[e.key].M||"–"}</td>,
                          <td key={e.key+"f"} style={{...td,textAlign:"center",background:f.counts[e.key].F>0?e.bg:"",color:e.color,fontWeight:f.counts[e.key].F>0?700:400}}>{f.counts[e.key].F||"–"}</td>
                        ])}
                      </tr>
                    ))}
                    <tr style={{background:"#1e3a5f",color:"#fff",fontWeight:700}}>
                      <td style={{...td,color:"#fff"}} colSpan={2}>TOTALES</td>
                      <td style={{...td,textAlign:"center",color:"#7dd3fc"}}>{totales.M}</td>
                      <td style={{...td,textAlign:"center",color:"#f9a8d4"}}>{totales.F}</td>
                      <td style={{...td,textAlign:"center",color:"#fff",fontSize:14}}>{totales.total}</td>
                      {ESTADOS.map(e=>[
                        <td key={e.key+"m"} style={{...td,textAlign:"center",color:"#7dd3fc",fontWeight:700}}>{totE[e.key].M||"–"}</td>,
                        <td key={e.key+"f"} style={{...td,textAlign:"center",color:"#f9a8d4",fontWeight:700}}>{totE[e.key].F||"–"}</td>
                      ])}
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          );
        })()}

        {tab==="reportes"&&(isAdmin||isCoord)&&(
          <ReportesTab datosEscuela={datosEscuela} appState={appState}/>
        )}

      </div>
    </div>
  );
}
