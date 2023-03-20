package main

import (
	"fmt"
	"log"
	"net/http"
	"time"

	"github.com/gorilla/mux"
	"github.com/xuri/excelize/v2"
	"gorm.io/driver/mysql"
	"gorm.io/gorm"
)

func main() {
	router := mux.NewRouter().StrictSlash(true)
	router.HandleFunc("/", DownloadExcel)
	log.Fatal(http.ListenAndServe(":9999", router))
}

type Data struct {
	ID       int    `json:"id"`
	Name     string `json:"name"`
	Email    string `json:"email"`
	Password string `json:"password"`
}

func connectDB() *gorm.DB {
	db, err := gorm.Open(mysql.Open("zidane:zidane@tcp(localhost:3306)/golangexcel?charset=utf8&parseTime=True&loc=Local"), &gorm.Config{})
	if err != nil {
		panic("failed to connect database")
	}
	return db

}

func (Data) TableName() string {
	return "users"
}

func DownloadExcel(w http.ResponseWriter, r *http.Request) {
	f := excelize.NewFile()
	db := connectDB()
	var data []Data
	db.Find(&data)

	// create title for excel
	f.SetCellValue("Sheet1", "A1", "Nomor")
	f.SetCellValue("Sheet1", "B1", "Name")
	f.SetCellValue("Sheet1", "C1", "Email")
	f.SetCellValue("Sheet1", "D1", "Password")

	// create data for excel

	var startRow = 2

	for i, v := range data {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i+startRow), v.ID)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i+startRow), v.Name)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i+startRow), v.Email)
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", i+startRow), v.Password)
	}

	// auto set column depending on data length
	for letter := 'A'; letter <= 'D'; letter++ {
		f.SetColWidth("Sheet1", string(letter), string(letter), 20)
	}

	style, err := f.NewStyle(&excelize.Style{
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
	})
	if err != nil {
		fmt.Println(err)
	}

	// apply after startrow
	f.SetCellStyle("Sheet1", "A1", fmt.Sprintf("D%d", len(data)+1), style)

	// make this download able
	uuid := time.Now().Format("20060102150405")
	filename := "report-" + uuid + ".xlsx"

	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	w.Header().Set("Content-Type", r.Header.Get("Content-Type"))

	if err := f.Write(w); err != nil {
		fmt.Println(err)
	}

}
