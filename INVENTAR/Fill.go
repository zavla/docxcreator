// to debug run
// curl.exe -T .\ex1.txt --url http://127.0.0.1:1313/docxcreator --verbose --output .\out2.docx
package main

import (
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"path/filepath"

	"baliance.com/gooxml/document"
)

// zaHeaderAndTable то что передается нам в json
type zaHeaderAndTable struct {
	DocxTemplateName string
	Header           map[string]string
	Table1           []map[int]string
}

type serviceConfig struct {
	pathToTemplates string
	showFields      bool
	fullPathLogFile string
}

type responseStruct struct {
	Error string
	Data  []byte
}

var currentConfig serviceConfig
var logfile *log.Logger

func makeresponse(Data []byte, rerr string) []byte {
	resp := &responseStruct{
		Error: rerr,
		Data:  Data,
	}
	b, err := json.Marshal(&resp)
	if err != nil {
		logfile.Printf("%s", err)
	}
	return b
}

func main() {
	pathToTemplates := flag.String("PathToTemplates", "", "путь к файлам-шаблонам")
	jsonFileName := flag.String("jsonFileName", "", `имя файла с данными json. Пример: {"Header":{"key1":"val1"},"Table1":[{"keytab1":"valtab1"}]}`)
	showFields := flag.Bool("ПоказатьСписокПолейШаблона", false, "распечатать список merge-полей файла шаблона")
	bindAddressPort := flag.String("bindAddressPort", "127.0.0.1:8080", "слушать на адресе")
	//Ключ := flag.String("КлючУникальности", "", "добавка к имени результирующего файла")
	fullPathLogFile := flag.String("logfile", "", "полный путь к лог файлу")
	flag.Parse()

	if *fullPathLogFile == "" {
		flag.Usage()
		log.Fatalf("%s\n", "не передан параметр logfile - полный путь к лог файлу")
	}

	flog, err := os.OpenFile(*fullPathLogFile, os.O_RDWR|os.O_CREATE, 0)
	if err != nil {
		log.Fatalf("%s\n%s\n", "не смогло создать лог файл "+(*fullPathLogFile), err)
	}
	logfile = log.New(flog, "", log.Ldate|log.Ltime)

	if *pathToTemplates == "" {
		flag.Usage()
		log.Fatalf("%s\n", "не передан параметр PathToTemplates - путь к каталогу шаблонов документов")
	}
	// if *Ключ == "" {
	// 	flag.Usage()
	// 	log.Fatalf("%s\n", "не передан параметр КлючУникальности.")
	// }
	if *jsonFileName == "" { //не передан файл, запустить http

		// global var
		currentConfig = serviceConfig{
			pathToTemplates: *pathToTemplates,
			showFields:      *showFields,
			fullPathLogFile: *fullPathLogFile,
		}

		var hand http.HandlerFunc = handlerhttp
		// err = http.ListenAndServe("127.0.0.1:1313", hand)
		err = http.ListenAndServe(*bindAddressPort, hand)
		fmt.Printf("%s\n%s", "Http server Exited:", err)

	} else { //передан файл с данными
		f, err := os.Open(*jsonFileName)
		if err != nil {
			log.Fatalf("%s\n%s\n", err, "файл с данными не открывается.")
		}
		databytes, err := ioutil.ReadAll(f)

		//tempdir := os.TempDir()
		новфайл := "debugfile.docx" //сохранение имя
		newfilefullpath := filepath.Join(".\\", новфайл)
		wdebug, err := os.OpenFile(newfilefullpath, os.O_CREATE|os.O_WRONLY, 0) //for debug
		if err != nil {
			logfile.Printf("%s\n%s", err, "Ошибка: результат не записывается в файл")
			os.Exit(1)
		}

		err = CreateDocxFromStruct(wdebug, databytes, *pathToTemplates, *showFields)
	}

}

func handlerhttp(w http.ResponseWriter, r *http.Request) {
	//action(w, )
	if r.URL.Path == "/docxcreator" {
		if r.Method == "GET" {
			w.WriteHeader(http.StatusBadRequest)
			w.Write(makeresponse([]byte{}, fmt.Sprintf("%s", "POST should be used.")))
			return
		}
		rdr := r.Body

		//debug
		// jsonbytes := make([]byte, 0, 3000)
		// jsonbytes, err := ioutil.ReadAll(rdr)
		// if err != nil {
		// 	logfile.Printf("%s", err)
		// }
		// w.Write(jsonbytes)
		err := action(w, rdr)
		if err != nil {
			logfile.Printf("%s", err)
			w.WriteHeader(http.StatusBadRequest)
			w.Write(makeresponse([]byte{}, fmt.Sprintf("%s", err)))
		}

	}
}

func action(w io.Writer, toreadbytes io.ReadCloser) error {

	var databytes []byte //то что прочитано из файла
	databytes, err := ioutil.ReadAll(toreadbytes)
	err = CreateDocxFromStruct(w, databytes, currentConfig.pathToTemplates, currentConfig.showFields)
	if err != nil {
		logfile.Printf("%s\n%s\n", "Ошибка: при создании документа.", err)
		return err
	}
	return nil
}

// CreateDocxFromStruct creates doxc document through gooxml and fills mergefields and adds
// rows to table. Fills rows from databytes which are json utf8 encoded struct zaHeaderAndTable.
func CreateDocxFromStruct(w io.Writer, databytes []byte, pathToTemplates string, showFields bool) error {

	datastr, err := getzaHeaderAndTable(databytes) //converts json to struct
	if err != nil {
		logfile.Printf("%s\n%s\n", "json не разбирается", err)
		return err
	}
	//log.Printf("%v\n", datastr)

	if datastr.DocxTemplateName == "" {
		errstr := "в json не передано поле DocxTemplateName, имя файла-шаблона."
		logfile.Printf("%s\n", errstr)
		return errors.New(errstr)
	}

	//opens template
	doc, err := document.Open(filepath.Join(pathToTemplates, datastr.DocxTemplateName))
	if err != nil {
		flag.Usage()
		logfile.Printf("%s\n%s", err, "docx файл-шаблон не читается.")
		return err
	}
	// When Word saves a document, it removes all unused styles.  This means to
	// copy the styles from an existing document, you must first create a
	// document that contains text in each style of interest.  As an example,
	// see the template.docx in this directory.  It contains a paragraph set in
	// each style that Word supports by default.
	// for _, s := range doc.Styles.Styles() {
	// 	fmt.Println("style", s.Name(), "has ID of", s.StyleID(), "type is", s.Type())
	// }

	//дозаполнение непереданных но существующих полей
	for _, v := range doc.MergeFields() {
		if showFields {
			fmt.Fprintf(w, "%s\n", v)
		}
		_, is := datastr.Header[v]
		if !is {
			datastr.Header[v] = " <не вказано> "
		}
	}

	doc.MailMerge(datastr.Header) //вставка шапки в док

	//ЗАПОЛНЕНИЕ ТАБЛИЦЫ
	//найти нужную строку
	var tabfound bool
	var tabindex int
	var totalcells int //сколько возможно передеть колонок в таблицу

	tabfound, tabindex, totalcells = findOurTable(doc) //поиск таблицы

	if tabfound {
		tab := doc.Tables()[tabindex]
		tab.Properties().SetStyle("TableGridZa")                      //в исходной документе должен быть этот стиль "таблицы"
		addRowWithCellsAndFillTexts(&tab, totalcells, datastr.Table1) //добавление строк
	}

	//сохранение результата
	err = doc.Save(w)
	if err != nil {
		logfile.Printf("%s", "Ошибка: невозможна запись в io.writer")
		return err
	}
	return nil
}

func getzaHeaderAndTable(bstr []byte) (zaHeaderAndTable, error) {
	var v zaHeaderAndTable
	err := json.Unmarshal(bstr, &v)
	if err != nil {
		log.Printf("%s\n%s", "не смогло прочесть json", err)
		return v, err

	}
	return v, nil

}

func findOurTable(doc *document.Document) (bool, int, int) {
	//seeks the table with the row with cells with text
	//1 2
	var tabfound bool
	var tabindex int
	var totalcells int //сколько возможно передеть колонок в таблицу

	tables := doc.Tables()
	for i, tab := range tables {
		//col1text := tab.Rows()[0].Cells()[0].Paragraphs()[0].Runs()[0].Text()
		rows := tab.Rows()
		for _, row := range rows {
			if len(row.Cells()) < 2 {
				continue
			}
			cell0 := row.Cells()[0]
			cell1 := row.Cells()[1]
			if len(cell0.Paragraphs()) < 1 || len(cell1.Paragraphs()) < 1 {
				continue
			}
			if len(cell0.Paragraphs()[0].Runs()) < 1 || len(cell1.Paragraphs()[0].Runs()) < 1 {
				continue
			}
			col1text := cell0.Paragraphs()[0].Runs()[0].Text()
			col2text := cell1.Paragraphs()[0].Runs()[0].Text()
			if col1text == "1" && col2text == "2" {
				tabfound = true
				tabindex = i
				totalcells = len(row.Cells()) //сколько возможно передеть колонок в таблицу
				break
			}
			// 	fmt.Printf("%s\n", string(j))
			// 	for c, cell := range row.Cells() {
			// 		fmt.Printf("%s", string(c))
			// 		for _, par := range cell.Paragraphs() {
			// 			for _, run := range par.Runs() {
			// 				curtext := run.Text()
			// 				fmt.Printf("%s ", curtext)
			// 			}
			// 		}
			// 	}
		}

	}
	return tabfound, tabindex, totalcells

}

func addRowWithCellsAndFillTexts(tab *document.Table, totalcells int, sliceofmaps []map[int]string) {
	for _, datamap := range sliceofmaps {

		nrow := tab.AddRow()
		for nc := 1; nc <= totalcells; nc++ {

			ncell := nrow.AddCell()
			npar := ncell.AddParagraph()

			nrun := npar.AddRun()
			nrun.AddText(datamap[nc]) //нам передан номер колонки по порядку
		}

	}
}
