*** Settings ***
Library           SeleniumLibrary
Library           clipboard
Library           String
Library           DateTime
Library           ExcelLibrary
Library           Dialogs

*** Variables ***
${Url}            https://reformaprocesal.pjud.cl/ConsultaCausasJsfWeb/page/panelConsultaCausas.jsf    # Direccion de la pagina a realizar las consultas
${PathExcel}      resultado/Nombres.xls    #Ubicacion de archivo Excel a consultar.
${NombreHojaExcel}    nombres    #Nombre de la hoja excel que se consulta.
${ContadorTribunalOrigen}    0
${Contador}       1    #Contador que recorrera el total de valores de archivo excel.
${NombreCopiar}    ${EMPTY}    #Nombre que se extrae de excel
${ApellidoPaternoCopiar}    ${EMPTY}
${ApellidoMaternoCopiar}    ${EMPTY}
${RutCopiar}      ${EMPTY}
${YearCopiar}     ${EMPTY}
${ContadorCasos}    ${EMPTY}
${ContadorDeCasosInternos}    1
${SiTienenCaso}    ${EMPTY}
${CounterInside}    1
${Var12}          1
@{ParaGuardarEnExcelSumarizado}    # Listado Para guardar En excel
${ParaGuardarEnExcelSumarizado}    ${EMPTY}
${Col}            0
${Conta}          ${EMPTY}
@{TotalJuzgado}    6    14    26    27    34    47    51    53    88    94    103    114    119    127    132    140    150
...               151    152    157    158    159    160    188    189    190    193    195    206    207    209    214    216    223
...               225    226    227    240    241    244    245    248    249    250    257    258    385    386    659    660    662
...               927    928    929    930    931    932    933    934    935    936    937    938    939    940    941    942    943
...               944    946    947    948    949    950    951    952    953    954    955    956    957    958    959    960    961
...               962    963    964    965    966    967    968    988    989    990    991    992    993    994    995    996    1013
...               1045    1046    1047    1048    1049    1050    1051    1052    1053    1054    1055    1056    1057    1058    1059    1060    1061
...               1062    1063    1064    1065    1067    1068    1069    1070    1071    1072    1073    1074    1075    1076    1077    1078    1079
...               1080    1081    1082    1083    1084    1085    1086    1087    1088    1089    1090    1091    1092    1093    1094    1095    1096
...               1097    1150    1151    1152    1220    1221    1222    1223    1224    1225    1226    1227    1228    1229    1230    1231    1232
...               1233    1234    1235    1236    1237    1238    1239    1240    1244    1245    1246    1247    1248    1249    1250    1251    1320
...               1321    1322    1323    1325    1326    1328    1355    1500    1501    1502
@{TotalJuzgado1}    1502    6    1501

*** Test Cases ***
TestFinal
    Open Excel    ${PathExcel}
    ${Count1}    Get Row Count    ${NombreHojaExcel}    #Total de filas
    @{Count1}    Get column values    ${NombreHojaExcel}    1    #Valores de la columna 1
    FOR    ${Var1}    IN    @{Count1}    #Recorre    cada fila de archivo excel
        BuscadorDeCasos1
        log    ${Contador}
        Sleep    5s
        AumentadorDeNumeroPorCaso
        Log    ${Var1}
    END
    #Log List    ${ParaGuardarEnExcelSumarizado}    1
    #log    Get Index From List    ${ParaGuardarEnExcelSumarizado}    1
    #Open Excel    resultado/Prototipo.xls
    #Put String To Cell    resultado    0    ${Contador}    ${RutCopiar}
    #FOR    ${var}    IN    ${ParaGuardarEnExcelSumarizado}
    #    Put String To Cell    resultado    1    1    ${var}${\n}
    #END
    #sleep    8s
    #${timestamp} =    Get Current Date    result_format=%Y-%m-%d-%H-%M
    #${filename} =    Set Variable    resultado-${timestamp}.xls
    #Save Excel    resultado/${filename}
    #Create File    resultado/resultado.txt
    #Append To File    resultado/resultado.txt    ${ParaGuardarEnExcelSumarizado}
    #Append To File    resultado/resultado.txt    ${ParaGuardarEnExcelSumarizado}    ${\n}
    #FOR    ${vAR}    IN    ${ParaGuardarEnExcelSumarizado}
    #Append To Csv File    resultado/resultado.cvs    ${ParaGuardarEnExcelSumarizado}    ${\n}
    #FOR    ${vAR}    IN    ${ParaGuardarEnExcelSumarizado}
    #    Append To File    resultado/resultado.txt    ${vAR}    \n
    #END

Reforma
    Open Browser    https://reformaprocesal.pjud.cl/ConsultaCausasJsfWeb/page/panelConsultaCausas.jsf    chrome
    Click Element    //*[@id="formConsultaCausas:idTabNombre_lbl"]
    Sleep    5s
    Close Browser

Check created excel doc
    Open Excel    ${PathExcel}
    ${Count1}    Get Row Count    ${NombreHojaExcel}    #Total de filas
    @{Count1}    Get column values    ${NombreHojaExcel}    1    #Valores de la columna 1
    FOR    ${Var1}    IN    @{Count1}    #Recorre    cada fila de archivo excel
        BuscadorDeCasos
        log    ${Contador}
    #Contador
        Sleep    5s
    #ValidarTotalCasos
    #ContadorCasosInternosReset
        Close Browser
    #AumentadorDeNumeroPorCaso
        Log    ${Var1}
    END

*** Keywords ***
BuscadorDeCasos
    [Documentation]    Rescata variables desde Excel.
    Open Excel    ${PathExcel}
    Open Browser    ${Url}    chrome    #Apertura de explorador
    Sleep    10s    \    #Espera de 10 segundos
    #Select Frame    name=formConsultaCausas
    Click Element    //*[@id="formConsultaCausas:idTabNombre_lbl"]
    log    ${Contador}
    ${NombreCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    B${Contador}
    clipboard.Copy    ${NombreCopiar}
    ${NombreCopiar}    Set Suite Variable    ${NombreCopiar}
    Log    ${NombreCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormNombres"]
    Press Keys    none    CTRL+V
    ${ApellidoPaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    D${Contador}
    clipboard.Copy    ${ApellidoPaternoCopiar}
    ${ApellidoPaternoCopiar}    Set Suite Variable    ${ApellidoPaternoCopiar}
    Log    ${ApellidoPaternoCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormApPater"]
    Press Keys    none    CTRL+V
    ${ApellidoMaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    E${Contador}
    clipboard.Copy    ${ApellidoMaternoCopiar}
    ${ApellidoMaternoCopiar}    Set Suite Variable    ${ApellidoMaternoCopiar}
    Log    ${ApellidoMaternoCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormApMater"]
    Press Keys    none    CTRL+V
    ${RutCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    A${Contador}
    clipboard.Copy    ${RutCopiar}
    ${RutCopiar}    Set Suite Variable    ${RutCopiar}
    Log    ${RutCopiar}
    ${YearCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    F${Contador}
    clipboard.Copy    ${YearCopiar}
    ${YearCopiar}    Set Suite Variable    ${YearCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormFecEra"]
    Press Keys    none    CTRL+V
    Log    ${YearCopiar}
    Sleep    10s
    Press Keys    \    ENTER
    Sleep    10s
    Click Element    //*[@id="formConsultaCausas:buscar2"]
    Sleep    18s
    FOR    ${VAR}    IN    @{TotalJuzgado1}
        Sleep    15s
        Select From List By Value    name:COD_TribunalSinTodos    ${VAR}
        Sleep    5s
        log    ${VAR}
        Click Element    //*[@id="formConsultaCausas:buscar2"]
        Sleep    10s    #${ContadorCasoExiste}=    Get Element Count    (//a[@onclick='ValDobleSubmit()'])[1]    #Sleep    10s    #Run Keyword If    ${ContadorCasoExiste}>0    RecorrerCasosInternos    #...    # ELSE    log    "No existe Registro Valido"
    #Sleep    10s
    END
    Close Browser

BuscadorDeCasos1
    [Documentation]    Rescata variables desde Excel.
    Open Excel    ${PathExcel}
    Open Browser    ${Url}    chrome    #Apertura de explorador
    Sleep    10s    \    #Espera de 10 segundos
    #Select Frame    name=formConsultaCausas
    Click Element    //*[@id="formConsultaCausas:idTabNombre_lbl"]
    log    ${Contador}
    ${NombreCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    B${Contador}
    clipboard.Copy    ${NombreCopiar}
    ${NombreCopiar}    Set Suite Variable    ${NombreCopiar}
    Log    ${NombreCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormNombres"]
    Press Keys    none    CTRL+V
    ${ApellidoPaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    D${Contador}
    clipboard.Copy    ${ApellidoPaternoCopiar}
    ${ApellidoPaternoCopiar}    Set Suite Variable    ${ApellidoPaternoCopiar}
    Log    ${ApellidoPaternoCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormApPater"]
    Press Keys    none    CTRL+V
    ${ApellidoMaternoCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    E${Contador}
    clipboard.Copy    ${ApellidoMaternoCopiar}
    ${ApellidoMaternoCopiar}    Set Suite Variable    ${ApellidoMaternoCopiar}
    Log    ${ApellidoMaternoCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormApMater"]
    Press Keys    none    CTRL+V
    ${RutCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    A${Contador}
    clipboard.Copy    ${RutCopiar}
    ${RutCopiar}    Set Suite Variable    ${RutCopiar}
    Log    ${RutCopiar}
    ${YearCopiar}    Read Cell Data By Name    ${NombreHojaExcel}    F${Contador}
    clipboard.Copy    ${YearCopiar}
    ${YearCopiar}    Set Suite Variable    ${YearCopiar}
    Click Element    //*[@id="formConsultaCausas:idFormFecEra"]
    Press Keys    none    CTRL+V
    Log    ${YearCopiar}
    #Click Element    //*[@id="formConsultaCausas:buscar2"]
    Sleep    18s
    FOR    ${VAR}    IN    @{TotalJuzgado1}
        Sleep    15s
        Select From List By Value    //*[@id="formConsultaCausas:idSelectedCodeTribunalNom"]    ${VAR}
        Sleep    5s
        log    ${VAR}
        Execute Manual Step    Please complete the CAPTCHA portion of the form.
        Sleep    1s
        Click Element    //*[@id="formConsultaCausas:buscar2"]
        Sleep    10s    #${ContadorCasoExiste}=    Get Element Count    (//a[@onclick='ValDobleSubmit()'])[1]    #Sleep    10s    #Run Keyword If    ${ContadorCasoExiste}>0    "Hay Caso"    # ELSE    log    "No existe Registro Valido"
        Sleep    10s
    END
    Close Browser

AumentadorDeNumeroPorCaso
    [Documentation]    Contador del total de personas de los cuales se consideraran para las consultas.
    ${temp}    Evaluate    ${Contador} + 1
    Set Test Variable    ${Contador}    ${temp}
