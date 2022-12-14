# СУБД 2022
**<div align="center">Лабораторная работа по дисциплине "Системы управления базами данных"</div>**

**<div align="center">Вариант 5. Управление организацией экспертизы научно-технических проектов.</div>**

**Объект разработки** – автоматизированное рабочее место организаторов экспертизы научно-технических проектов.

**Цель работы**: создание средств поддержки формирования экспертных групп для проведения экспертизы научно-технических проектов.

**Информационный базис**: данные об ученых, давших принципиальное согласие на участие в экспертизах по своей области интересов.

**Экспертиза научно-технического проекта** (НТП) - формирование аргументированного заключения о целесообразности финансирования НТП на основе анализа его актуальности, научной состоятельности, технико-экономических характеристик. Заключение формируется на основе мнений нескольких независимых экспертов, составляющих экспертную группу по конкретному НТП. В начале экспертизы группа должна быть выбрана в соответствии с предметной областью НТП и областями интересов экспертов.

Структура информации (файл _EXPERT.XLS_): код эксперта (символьный), фамилия и.о. эксперта, регион проживания, город, код ГРНТИ области интересов, ключевые слова, характеризующие область интересов, число участий в экспертизах, дата занесения в базу данных.

Структура информации о рубриках ГРНТИ (файл _GRNTIRUB.XLS_): наименование рубрики, код рубрики.

Справочная таблица _REG_OBL_CITY.XLS_ вхождения субъектов федерации (_oblname_) и городов (_city_) в федеральные округа (_region_).

**Требования к функциям**, реализуемым в программах анализа данных:

* контроль и восстановление целостности исходных баз данных системы;

* добавление/редактирование информации об эксперте в базу данных, верификация вновь поступивших данных, обеспечение целостности данных;

* контроль возможного повторного занесения данных об эксперте;

* фильтрация информации в базе данных по указанной фамилии и/или федеральному округу и/или субъекту федерации и/или городу и/или рубрике или коду ГРНТИ и/или ключевым словам области интересов; фиксация отобранного подмножества в поименованную экспертную группу;

* просмотр записей выбранной группы кандидатов на включение в состав экспертной группы с возможностью простановки/снятия отметок о принятии решения о включении кандидата в экспертную группу; фиксация результата в экспертной группе;

* просмотр записей исходной базы данных с возможностью простановки отметок об отборе эксперта в качестве кандидата на включение/добавление в экспертную группу, перенос сведений об отобранных кандидатах в выбранную экспертную группу;

* утверждение экспертной группы без возможности дальнейшей корректировки состава с увеличением на 1 числа участий в экспертизах в основной базе данных;

* формирование документов: таблица со списком сформированной поименованной экспертной группы, содержащей столбцы: порядковый номер, фамилия И.О., регион, город, код ГРНТИ; карточка эксперта.
