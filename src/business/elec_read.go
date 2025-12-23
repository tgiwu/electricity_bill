package business

import (
	"electricity_bill/src/types"
	"electricity_bill/src/utils"
	"fmt"
	"log"
	"reflect"
	"strconv"
	"strings"

	"github.com/spf13/viper"
	"github.com/tealeg/xlsx/v3"
)

const (
	ELEC_CONTENT_ROOM_NO     = "门牌号"
	ELEC_CONTENT_TIMES       = "倍率"
	ELEC_CONTENT_ORIGINAL    = "表底"
	ELEC_CONTENT_INDICATION  = "度数"
	ELEC_CONTENT_AIR_CONTRAL = "空调"
	ELEC_CONTENT_COST        = "用电量"
	ELEC_CONTENT_ALL_COST    = "总用电量"

	// ELEC_COLS_PER_MONTH = 4
)

// 与月份无关的独立列数
var rowNoDataStart = 0

func ReadElec(ce *chan types.Indication, finish *chan string) error {

	file, err := xlsx.OpenFile(viper.GetString("elec_file"))

	if err != nil {
		return err
	}

	readSheets(file, ce, finish)
	*finish <- "ele_f"
	return nil
}

func readSheets(file *xlsx.File, ce *chan types.Indication, finish *chan string) error {

	if sheet, found := file.Sheet[viper.GetString("indication_sheet")]; found {
		readSheetIndic(sheet, ce, finish)
	} else {
		log.Panic("can not find indication sheet !!!!")
	}
	return nil
}

func readSheetIndic(sheet *xlsx.Sheet, ce *chan types.Indication, finish *chan string) error {
	headerList := make(map[int]string, sheet.MaxCol)

	err := readIndicHeader(sheet, &headerList, finish)

	if err != nil {
		return err
	}
	// fmt.Printf("%s : %v \n", sheet.Name, headerList)

	err = readElecData(*sheet, &headerList, ce, finish)

	fmt.Println("elec read done ")
	fmt.Println("-------------")
	return err
}

func readIndicHeader(sheet *xlsx.Sheet, headerList *map[int]string, finish *chan string) error {

	targetMonth := viper.GetInt("target_month")

	for rowIndex := range sheet.MaxRow {
		row, err := sheet.Row(rowIndex)
		if err != nil {
			panic(err)
		}

		//find header row by cell value CONTENT_ROOM_NO
		if v := strings.TrimSpace(row.GetCell(0).Value); v == CONTENT_ROOM_NO {
			log.Println("header row is ", rowIndex)
			(*headerList)[0] = "RoomNo"
			var (
				targetDataIndicLastMonth string
				targetDataIndic          = fmt.Sprintf("%d-%s", targetMonth, ELEC_CONTENT_INDICATION)
				targetDataAirControl     = fmt.Sprintf("%d-%s", targetMonth, ELEC_CONTENT_AIR_CONTRAL)
				targetDataCost           = fmt.Sprintf("%d-%s", targetMonth, ELEC_CONTENT_COST)
				targetDataAllCost        = fmt.Sprintf("%d-%s", targetMonth, ELEC_CONTENT_ALL_COST)
			)

			//first month
			if targetMonth == 1 {
				targetDataIndicLastMonth = ELEC_CONTENT_ORIGINAL
			} else {
				targetDataIndicLastMonth = fmt.Sprintf("%d-%s", targetMonth-1, ELEC_CONTENT_INDICATION)
			}

			for colIndex := 1; colIndex < sheet.MaxCol; colIndex++ {
				v = strings.TrimSpace(row.GetCell(colIndex).Value)
				if len(v) == 0 {
					continue
				}
				switch v {
				case ELEC_CONTENT_TIMES:
					(*headerList)[colIndex] = "Times"
				case targetDataAirControl:
					(*headerList)[colIndex] = "CostAirControal"
				case targetDataIndicLastMonth:
					(*headerList)[colIndex] = "IndicLastMonth"
				case targetDataIndic:
					(*headerList)[colIndex] = "Indic"
				case targetDataCost:
					(*headerList)[colIndex] = "Cost"
				case targetDataAllCost:
					(*headerList)[colIndex] = "CostAll"
				default:
					//ignore
				}
			}
			//data start row
			rowNoDataStart = rowIndex + 1
			break
		} else {
			continue
		}
	}

	*finish <- "header read finish"

	return nil
}

func readElecData(sheet xlsx.Sheet, headers *map[int]string, ce *chan types.Indication, finish *chan string) error {
	rowCount := sheet.MaxRow
	for rowIndex := rowNoDataStart; rowIndex < rowCount; rowIndex++ {
		row, err := sheet.Row(rowIndex)
		if err != nil {
			return err
		}

		indication := types.Indication{}

		for index, content := range *headers {
			v := row.GetCell(index).Value

			field, _ := reflect.TypeOf(indication).FieldByName(content)
			value := reflect.ValueOf(&indication).Elem().FieldByName(content)
			switch field.Type.Kind() {
			case reflect.Float64:
				if len(strings.TrimSpace(v)) == 0 {
					value.SetFloat(0)
					continue
				}
				val, err := strconv.ParseFloat(strings.TrimSpace(v), 64)
				if err != nil {
					fmt.Println(err)
					value.SetFloat(0)
				} else {
					value.SetFloat(val)
				}
			case reflect.String:
				value.SetString(v)
			}

		}

		if len(indication.RoomNo) != 0 {
			err := utils.FindUnitFromRoomNo(indication.RoomNo, &indication.Unit, &indication.Floor)
			if err != nil {
				log.Panic(err)
			}
			*ce <- indication
		}
	}
	*finish <- "elec indication read finish"
	return nil
}
