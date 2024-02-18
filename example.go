package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

type Bill struct {
	Price int
}

type Address struct {
	Street      string
	HouseNumber int
	Bills       []Bill
}

type Person struct {
	Name      string
	Age       int
	Sex       string
	Addresses []Address
}

func main() {
	xxx := []Person{
		{
			Name: "Vasya",
			Age:  10, Sex: "M",
			Addresses: []Address{{Street: "Pushkina", HouseNumber: 1}, {Street: "Pushkina2", HouseNumber: 2}}},
		{Name: "Sergey", Age: 10, Sex: "M"},
	}
	x, y := GetFieldAxis(xxx)
	fmt.Println(x, y)
	fields := make(map[string][]any)
	fields = GetMap(xxx, fields)
	xlsx := excelize.NewFile()
	err := MakeSheetFMap(xlsx, fields, nil, "Sheet1", true, true)
	if err != nil {
		return
	}
	xlsx.SaveAs("TrahBabah.xlsx")
}
