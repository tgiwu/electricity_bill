package word

import (
	"electricity_bill/src/types"
	"electricity_bill/src/utils"
	"fmt"
	"log"
	"path"
	"strconv"
	"sync"

	"github.com/ZeroHawkeye/wordZero/pkg/document"
	"github.com/ZeroHawkeye/wordZero/pkg/style"
	"github.com/spf13/viper"
)

const (
	STYLE_SU_LITTLE_FOUR_CENTER_B = "STYLE_MAIN_TITLE"
	STYLE_SU_FIVE_LEFT_U          = "STYLE_SU_FIVE_LEFT_U"
	STYLE_SU_FIVE_LEFT            = "STYLE_SU_FIVE_LEFT"
	STYLE_SU_THREE_CENTER_B       = "STYLE_SU_THREE_CENTER_B"

	STYLE_TABLE_BILL = iota
	STYLE_TABLE_BILL_NO_PAY
	STYLE_TABLE_AIR_CONTROL
)

var (
	styleTableBillNormal = document.TableConfig{
		Cols:      6,
		Rows:      2,
		Width:     7500,
		ColWidths: []int{1000, 1500, 1000, 1500, 1500, 1000},
	}

	styleTableBillNoPay = document.TableConfig{
		Cols:      5,
		Rows:      2,
		Width:     7500,
		ColWidths: []int{1200, 1700, 1200, 1200, 1700},
	}

	styleTableBillAirControl = document.TableConfig{
		Cols:      2,
		Rows:      2,
		Width:     7500,
		ColWidths: []int{1500, 6000},
	}
)

var CN_NUMBER = [10]string{"零", "一", "二", "三", "四", "五", "六", "七", "八", "九"}

func setUpStyle(doc *document.Document) {
	styleManager := doc.GetStyleManager()
	quickAPI := style.NewQuickStyleAPI(styleManager)

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   STYLE_SU_LITTLE_FOUR_CENTER_B,
		Name: STYLE_SU_LITTLE_FOUR_CENTER_B,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "center",
			SpaceAfter:  0,
			SpaceBefore: 0,
			LineSpacing: 0,
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体 (中文正文)",
			FontSize:  12,
			FontColor: "000000",
			Bold:      true,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   STYLE_SU_FIVE_LEFT,
		Name: STYLE_SU_FIVE_LEFT,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0,
			SpaceAfter:  0,
			LineSpacing: 0,
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体 (中文正文)",
			FontSize:  11,
			FontColor: "000000",
			Bold:      false,
			Underline: false,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   STYLE_SU_FIVE_LEFT_U,
		Name: STYLE_SU_FIVE_LEFT_U,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "left",
			SpaceBefore: 0,
			SpaceAfter:  0,
			LineSpacing: 0,
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体 (中文正文)",
			FontSize:  11,
			FontColor: "000000",
			Bold:      false,
			Underline: true,
		},
	})

	quickAPI.CreateQuickStyle(style.QuickStyleConfig{
		ID:   STYLE_SU_THREE_CENTER_B,
		Name: STYLE_SU_THREE_CENTER_B,
		Type: style.StyleTypeParagraph,
		ParagraphConfig: &style.QuickParagraphConfig{
			Alignment:   "center",
			SpaceBefore: 0,
			SpaceAfter:  0,
			LineSpacing: 0,
		},
		RunConfig: &style.QuickRunConfig{
			FontName:  "宋体",
			FontSize:  16,
			FontColor: "000000",
			Bold:      true,
			Underline: false,
		},
	})
}

func CreateDocxs(indicMap *map[int]map[string]types.Indication, companiesMap *map[int]map[string]types.CompanyInfo, finish *chan string) {

	var wg sync.WaitGroup
	count := 0

	for unit, companies := range *companiesMap {
		if indics, found := (*indicMap)[unit]; found {
			count++
			wg.Add(1)
			go createSingleDocx(&indics, &companies, unit, finish)
		}
	}
	wg.Wait()
	*finish <- "All docxs create finish !!"
}

func createSingleDocx(indics *map[string]types.Indication, companies *map[string]types.CompanyInfo, unit int, finish *chan string) {
	doc := document.New()
	setUpStyle(doc)
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationPortrait)
	doc.SetPageMargins(5.56, 15, 5.56, 15)

	for _, info := range *companies {
		if indic, found := (*indics)[info.GateNo]; found && info.IsNeedBill {
			createDocxPage(doc, &indic, &info)
		}

	}

	doc.Save(path.Join(viper.GetString("output"), fmt.Sprintf("%s-%d电费通知单-%s单元.docx", viper.GetString("target_year"), viper.GetInt("target_month"), CN_NUMBER[unit])))
	
	*finish <- fmt.Sprintf("unit %d finish\n", unit)
}

func createDocxPage(doc *document.Document, indic *types.Indication, companyInfo *types.CompanyInfo) {
	title(doc)
	floor(doc, companyInfo)
	nameAndAddress(doc, companyInfo)
	doc.AddParagraph("/n/n")
	sign(doc)
	doc.AddParagraph("/n")
	//table
	billInfo(doc, indic)
	//expense
	expense(doc, indic)
	//backup
	backup(doc)

}

func title(doc *document.Document) {
	para := doc.AddParagraph("电费缴费通知单")
	para.SetStyle(STYLE_SU_LITTLE_FOUR_CENTER_B)
}

func floor(doc *document.Document, info *types.CompanyInfo) {

	para := doc.AddParagraph(fmt.Sprintf("(户号：%s单元%s层)", CN_NUMBER[(*info).Unit], CN_NUMBER[(*info).Floor]))
	para.SetStyle(STYLE_SU_FIVE_LEFT_U)
}

func nameAndAddress(doc *document.Document, info *types.CompanyInfo) {
	paraName := doc.AddParagraph(fmt.Sprintf("客户名称：%s", info.Name))
	paraName.SetStyle(STYLE_SU_FIVE_LEFT)

	paraAddress := doc.AddParagraph(fmt.Sprintf("地    址：中国影都文娱产业园%s单元%s层", CN_NUMBER[info.Unit], CN_NUMBER[info.Floor]))
	paraAddress.SetStyle(STYLE_SU_FIVE_LEFT)
}

func sign(doc *document.Document) {
	para := doc.AddParagraph("缴费人确认：________________________")
	para.SetStyle(STYLE_SU_FIVE_LEFT)
}

func tableArea(doc *document.Document, indic *types.Indication, style int) {
	var config document.TableConfig = document.TableConfig{}
	tableTitle(doc, indic, style)
	tableConfig(indic, &config, style)
	if len(config.Data) == 0 {
		log.Fatal("unsupport table style ", style)
		return
	}
	doc.AddTable(&config)
}
func tableTitle(doc *document.Document, indic *types.Indication, style int) {
	var paraTitle = &document.Paragraph{}
	switch style {
	case STYLE_TABLE_BILL:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%s年%d月%s用电量", viper.GetString("target_year"), viper.GetInt("target_month"), indic.RoomNo))
	case STYLE_TABLE_BILL_NO_PAY:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%s年%d月%s用电量", viper.GetString("target_year"), viper.GetInt("target_month"), indic.RoomNo))
	case STYLE_TABLE_AIR_CONTROL:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%s年%d月%s外机空调用电量", viper.GetString("target_year"), viper.GetInt("target_month"), indic.RoomNo))
	}
	paraTitle.SetStyle(STYLE_SU_THREE_CENTER_B)

}

func tableConfig(indic *types.Indication, config *document.TableConfig, style int) {
	switch style {
	case STYLE_TABLE_BILL:
		config.Cols = styleTableBillNormal.Cols
		config.Rows = styleTableBillNormal.Rows
		config.Width = styleTableBillNormal.Width
		config.ColWidths = styleTableBillNormal.ColWidths
		config.Data = [][]string{
				{"月份", "上月表数", "倍率", "本月表数", "实际用量（度）", "应缴电费"},
				{fmt.Sprint(viper.GetInt("target_month")),
					fmt.Sprint(indic.IndicLastMonth),
					strconv.FormatFloat(indic.Times, 'f', 0, 64),
					strconv.FormatFloat(indic.Indic, 'f', 2, 64),
					strconv.FormatFloat(indic.Cost, 'f', 2, 64),
					strconv.FormatFloat(indic.Cost, 'f', 2, 64)},
			}

	case STYLE_TABLE_BILL_NO_PAY:
		config.Cols = styleTableBillNoPay.Cols
		config.Rows = styleTableBillNoPay.Rows
		config.Width = styleTableBillNoPay.Width
		config.ColWidths = styleTableBillNoPay.ColWidths
		config.Data = [][]string{
				{"月份", "上月表数", "倍率", "本月表数", "实际用量（度）"},
				{fmt.Sprint(viper.GetInt("target_month")),
					fmt.Sprint(indic.IndicLastMonth),
					strconv.FormatFloat(indic.Times, 'f', 0, 64),
					strconv.FormatFloat(indic.Indic, 'f', 2, 64),
					strconv.FormatFloat(indic.Cost, 'f', 2, 64)},
			}
	case STYLE_TABLE_AIR_CONTROL:
		config.Cols = styleTableBillAirControl.Cols
		config.Rows = styleTableBillAirControl.Rows
		config.Width = styleTableBillAirControl.Width
		config.ColWidths = styleTableBillAirControl.ColWidths
		config.Data = [][]string{
				{"月份", "实际用量（度）"},
				{fmt.Sprint(viper.GetInt("target_month")), strconv.FormatFloat(indic.CostAirControal, 'f', 2, 64)},
			}
	default:
		//ignore
	}
}

func billInfo(doc *document.Document, indic *types.Indication) {
	para := doc.AddParagraph("缴费信息")
	para.SetStyle(STYLE_SU_FIVE_LEFT)

	doc.AddParagraph("\n")

	if indic.RoomNo == "1-1-总" {
		tableArea(doc, indic, STYLE_TABLE_BILL)

	} else {
		tableArea(doc, indic, STYLE_TABLE_BILL_NO_PAY)
	}

	if indic.CostAirControal != 0 {
		tableArea(doc, indic, STYLE_TABLE_AIR_CONTROL)
	}
}

func expense(doc *document.Document, indic *types.Indication) {
	year, _ := strconv.Atoi(viper.GetString("target_year"))
	month := viper.GetInt("target_month")
	lastDayInMonth := utils.DaysInMonth(year, month)
	//charging cycles
	chargingCyclesPara := doc.AddParagraph(
		fmt.Sprintf("本期电费周期：%d年%d月1日 至 %d年%d月%d日（1个月）", year, month, year, month, lastDayInMonth))
	chargingCyclesPara.SetStyle(STYLE_SU_FIVE_LEFT)
	//electricity cost sum

	//electricity price

	//electricity pay

	//liquidated damages

	//expense sum

	//account info
}

func backup(doc *document.Document) {

	para := doc.AddParagraph(fmt.Sprintf("缴费截止日期\n请在 %s年%d月15日前缴纳，逾期将按日收取违约金（0.05%%-0.1%%/天）。", viper.GetString("target_year"), viper.GetInt("target_month")))
	para.SetStyle(STYLE_SU_FIVE_LEFT)
}
