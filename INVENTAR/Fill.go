package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"os"

	"baliance.com/gooxml/document"
)

type zaHeaderAndTable struct {
	Header map[string]string
	Table1 []map[int]string
}

func main() {
	docxFileName := flag.String("docxFileName", "", "имя файла-шаблона")
	jsonFileName := flag.String("jsonFileName", "", `имя файла с данными json. Пример: {"Header":{"key1":"val1"},"Table1":[{"keytab1":"valtab1"}]}`)
	Ключ := flag.String("КлючУникальности", "", "добавка к имени результирующего файла")
	flag.Parse()

	if *docxFileName == "" {
		log.Fatalf("не передан параметр docxFileName")
	}
	if *Ключ == "" {
		log.Fatalf("не передан параметр КлючУникальности")
	}

	var datastr zaHeaderAndTable //точто загружается в док
	var databytes []byte         //то что прочитано из файла

	if *jsonFileName == "" {
		//test values
		datastr.Header = map[string]string{
			"ФирмаНаименование": "ООО ХОЛОД",
			//"":"",
			"е1":            "1",
			"е2":            "2",
			"е3":            "3",
			"ДатаСкладання": "18-11-2018р.",
		}
		datastr.Table1 = []map[int]string{
			{1: "1", 2: "Товар1", 3: "год2018"},
		}
		bstr, err := json.Marshal(datastr)
		if err != nil {
			log.Fatalf("%s", "bad")
		}
		databytes = []byte(bstr)
		fmt.Printf("TEST DATA: %#v", datastr)
	} else {
		f, err := os.Open(*jsonFileName)
		if err != nil {
			log.Fatalf("%s\n%s", err, "файл с данными не открывается.")
		}
		databytes, err = ioutil.ReadAll(f)
		//reads json

	}
	datastr = getzaHeaderAndTable(databytes)
	log.Printf("%v", datastr)

	//

	//fills the doc
	doc, err := document.Open(*docxFileName)
	if err != nil {
		log.Fatalf("%s\n%s", err, "docx файл не читается.")
	}
	//отладка полей
	for _, v := range doc.MergeFields() {
		log.Printf("%s\n", v)
		_, is := datastr.Header[v]
		if !is {
			datastr.Header[v] = " <не вказано> "
		}
	}
	//незаполненные поля сделать ___________

	doc.MailMerge(datastr.Header) //вставка в док

	//найти нужную строку
	var tabfound bool = false
	var tabindex int
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
	if tabfound {
		tab := doc.Tables()[tabindex]
		for _, datamap := range datastr.Table1 {

			nrow := tab.AddRow()
			for c, cell := range nrow.Cells() {
				cell.AddParagraph().AddRun().AddText(datamap[c])
			}
		}
	}
	новфайл := string(*Ключ) + "=" + *docxFileName //сохранение
	err = doc.SaveToFile(новфайл)
	if err != nil {
		log.Fatalf("%s\n%s", err, "результат не записывается в файл")
	}
}

func getzaHeaderAndTable(bstr []byte) zaHeaderAndTable {
	var v zaHeaderAndTable
	err := json.Unmarshal(bstr, &v)
	if err != nil {
		log.Fatalf("%s\n%s", "не смогло прочесть json", err)
	}
	return v

}
