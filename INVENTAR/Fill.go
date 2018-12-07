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
	Table1 []map[string]string
}

func main() {
	docxFileName := flag.String("docxFileName", "", "имя файла-шаблона")
	jsonFileName := flag.String("jsonFileName", "", `имя файла с данными json. Пример: {"Header":{"key1":"val1"},"Table1":[{"keytab1":"valtab1"}]}`)
	flag.Parse()

	if *docxFileName == "" {
		log.Fatalf("не передан параметр docxFileName")
	}

	var datastr zaHeaderAndTable //точто загружается в док
	var databytes []byte         //то что прочитано из файла

	if *jsonFileName == "" {
		//test values
		datastr.Header = map[string]string{"key1": "val1"}
		datastr.Table1 = []map[string]string{{"keytab1": "valtab1"}}
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
	doc.MailMerge(datastr.Header)

}

func getzaHeaderAndTable(bstr []byte) zaHeaderAndTable {
	var v zaHeaderAndTable
	err := json.Unmarshal(bstr, &v)
	if err != nil {
		log.Fatalf("%s\n%s", "не смогло прочесть json", err)
	}
	return v

}
