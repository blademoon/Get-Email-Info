# Get-Email-Info (альфа версия, возможна нестабильная работа, разработка прекращена).

## ***ВАЖНО! ЭТО ТЕСТОВАЯ ВЕРСИЯ РАЗРАБОТКА КОТОРОЙ ПРЕКРАЩЕНА!*** 
Тестирование проводилось на Microsoft Exchange Server 2010.
Небольшой скрипт упрощающий работу. Собирает необходимую для согласования заявки информацию непосредственно с сервера Microsoft Echange 2010.

## Минимальные требования
1. Microsoft Powershell 5.1
2. Модули Exchange и Microsoft Active Directory. 

## Как использовать
Необходимо произвести настройки в самом скрипте:    
Строка 5 - имя проверяемого почтового ящика    
Строка 6 - Путь к файлу с результатами работы скрипта.    
Строка 55 - Путь к отладочному файлу, содержащему список найденных в почтовом ящике папок.    
Строка 269 - Список групп членство в которых необходимо проверить.    

Сам скрипт запускается непосредственно на сервере Microsoft Exchange. Лучше создать отдельную папку, в которой будет хранится скрипт и результаты его работы.

По окончании работы скрипт формирует выходные файлы в формате txt содержащие собранную информацию.

## Если вы нашли ошибку в работе скрипта или коде:
  Сообщите мне об этом по электронной почте _blademoon@yandex.ru_.
  В поле "_Тема_" обзятельно укажите "_Get-Email-Info_"
