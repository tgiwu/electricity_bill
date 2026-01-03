package word

import (
	"electricity_bill/src/types"
	"electricity_bill/src/utils"
	"fmt"
	"io/fs"
	"log"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

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
	STYLE_TABLE_BILL_NO_PAYMENT
	STYLE_TABLE_AIR_CONTROL
	STYLE_TABLE_AIR_CONTROL_NO_PAYMENT
)

var styleTextFiveNormal = document.TextFormat{
	Bold:       false,
	Italic:     false,
	FontSize:   11,
	FontColor:  "000000",
	FontFamily: "宋体 (中文正文)",
}

var (
	styleTableBillNormal = document.TableConfig{
		Cols:      6,
		Rows:      2,
		Width:     9000,
		ColWidths: []int{1500, 1500, 1500, 1500, 1500, 1500},
	}

	styleTableBillNoPay = document.TableConfig{
		Cols:      5,
		Rows:      2,
		Width:     9000,
		ColWidths: []int{1800, 1800, 1800, 1800, 1800},
	}

	styleTableBillAirControl = document.TableConfig{
		Cols:      3,
		Rows:      2,
		Width:     9000,
		ColWidths: []int{3000, 3000, 3000},
	}

	styleTableBillAirControlNoBill = document.TableConfig{
		Cols:      2,
		Rows:      2,
		Width:     9000,
		ColWidths: []int{3000, 6000},
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

func CreateDocxs(indicMap *map[int]map[string]types.Indication, companiesMap *map[int]map[string]types.CompanyInfo, finish *chan string, wg *sync.WaitGroup) {

	path := viper.GetString("output")

	filepath.Walk(path, func(path string, info fs.FileInfo, err error) error {
		_, e := os.Stat(path)
		if os.IsNotExist(e) {
			return nil
		}
		if info != nil && strings.HasSuffix(info.Name(), ".docx") {
			log.Println("remove ", path)
			os.Remove(path)
		}
		return nil
	})

	log.Println("start create docx. ")
	for unit, companies := range *companiesMap {
		if indics, found := (*indicMap)[unit]; found {
			go createSingleDocx(&indics, &companies, unit, finish)
		} else {
			*finish <- fmt.Sprintf("no_data_unit_%d_f", unit)
		}
	}
	*finish <- "docx create finish"
}

func createSingleDocx(indics *map[string]types.Indication, companies *map[string]types.CompanyInfo, unit int, finish *chan string) {
	document.SetGlobalLevel(document.LogLevelError)
	doc := document.New()
	setUpStyle(doc)
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationPortrait)
	doc.SetPageMargins(23, 27, 23, 27)

	keys := sortByGateNo(companies)

	for index, key := range keys {
		info := (*companies)[key]
		if indic, found := (*indics)[info.GateNo]; found && info.IsNeedBill {
			//calculation payment
			if info.IsAddPayment {
				indic.Payment = indic.Cost * info.RateOfPay
				indic.AirControlPayment = indic.CostAirControal * info.RateOfPay
			}
			createDocxPage(doc, &indic, &info)
			if index != len(keys)-1 {
				//break
				doc.AddPageBreak()
			}
		}

	}

	doc.Save(path.Join(viper.GetString("output"), fmt.Sprintf("%d-%d电费通知单-%s单元.docx", viper.GetInt("target_year"), viper.GetInt("target_month"), CN_NUMBER[unit])))

	*finish <- fmt.Sprintf("_f_unit_%d", unit)
}

func createDocxPage(doc *document.Document, indic *types.Indication, companyInfo *types.CompanyInfo) {
	title(doc)
	doc.AddParagraph("")
	floor(doc, companyInfo)
	nameAndAddress(doc, companyInfo)
	doc.AddParagraph("")
	doc.AddParagraph("")
	sign(doc)
	doc.AddParagraph("")
	//table
	billInfo(doc, indic, companyInfo)
	//expense
	expense(doc, indic, companyInfo)
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

func tableArea(doc *document.Document, indic *types.Indication, compInfo *types.CompanyInfo, style int) {
	var config document.TableConfig = document.TableConfig{}
	tableTitle(doc, compInfo, style)
	tableConfig(indic, &config, style)
	if len(config.Data) == 0 {
		log.Fatal("unsupport table style ", style)
		return
	}
	table, _ := doc.AddTable(&config)

	for row := 0; row < table.GetRowCount(); row++ {
		table.SetRowHeight(row, &document.RowHeightConfig{
			Height: 26,
		})
		for col := 0; col < table.GetColumnCount(); col++ {
			f, _ := table.GetCellFormat(row, col)
			f.HorizontalAlign = document.CellAlignCenter
			f.VerticalAlign = document.CellVAlignCenter
			table.SetCellFormat(row, col, f)
		}
	}
}
func tableTitle(doc *document.Document, companyInfo *types.CompanyInfo, style int) {
	var paraTitle = &document.Paragraph{}
	var room string
	if strings.HasSuffix(companyInfo.GateNo, "总") {
		room = fmt.Sprintf("%d单元%d层", companyInfo.Unit, companyInfo.Floor)
	} else {
		room = companyInfo.GateNo
	}
	switch style {
	case STYLE_TABLE_BILL:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%d年%d月%s用电量", viper.GetInt("target_year"), viper.GetInt("target_month"), room))
	case STYLE_TABLE_BILL_NO_PAYMENT:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%d年%d月%s用电量", viper.GetInt("target_year"), viper.GetInt("target_month"), room))
	case STYLE_TABLE_AIR_CONTROL:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%d年%d月%s外机空调用电量", viper.GetInt("target_year"), viper.GetInt("target_month"), room))
	case STYLE_TABLE_AIR_CONTROL_NO_PAYMENT:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%d年%d月%s外机空调用电量", viper.GetInt("target_year"), viper.GetInt("target_month"), room))

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
				strconv.FormatFloat(indic.Payment, 'f', 2, 64)},
		}

	case STYLE_TABLE_BILL_NO_PAYMENT:
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
			{"月份", "实际用量（度）", "应缴电费"},
			{fmt.Sprint(viper.GetInt("target_month")),
				strconv.FormatFloat(indic.CostAirControal, 'f', 2, 64),
				strconv.FormatFloat(indic.AirControlPayment, 'f', 2, 64)},
		}
	case STYLE_TABLE_AIR_CONTROL_NO_PAYMENT:
		config.Cols = styleTableBillAirControlNoBill.Cols
		config.Rows = styleTableBillAirControlNoBill.Rows
		config.Width = styleTableBillAirControlNoBill.Width
		config.ColWidths = styleTableBillAirControlNoBill.ColWidths
		config.Data = [][]string{
			{"月份", "实际用量（度）"},
			{fmt.Sprint(viper.GetInt("target_month")), strconv.FormatFloat(indic.CostAirControal, 'f', 2, 64)},
		}
	default:
		//ignore
	}
}

func billInfo(doc *document.Document, indic *types.Indication, compInfo *types.CompanyInfo) {
	para := doc.AddParagraph("缴费信息")
	para.SetStyle(STYLE_SU_FIVE_LEFT)

	doc.AddParagraph("")

	if compInfo.IsAddPayment {
		tableArea(doc, indic, compInfo, STYLE_TABLE_BILL)
	} else {
		tableArea(doc, indic, compInfo, STYLE_TABLE_BILL_NO_PAYMENT)
	}
	if indic.CostAirControal == 0 {
		return
	}

	doc.AddParagraph("")

	if compInfo.IsAddPayment {
		tableArea(doc, indic, compInfo, STYLE_TABLE_AIR_CONTROL)
	} else {
		tableArea(doc, indic, compInfo, STYLE_TABLE_AIR_CONTROL_NO_PAYMENT)
	}
	doc.AddParagraph("")
}

func expense(doc *document.Document, indic *types.Indication, companyInfo *types.CompanyInfo) {
	year := viper.GetInt("target_year")
	month := viper.GetInt("target_month")
	lastDayInMonth := utils.DaysInMonth(year, month)
	//charging cycles
	doc.AddParagraph("")
	chargingCyclesPara := doc.AddParagraph(
		fmt.Sprintf("1.本期电费周期：%d年%d月1日 至 %d年%d月%d日（1个月）", year, month, year, month, lastDayInMonth))
	chargingCyclesPara.SetStyle(STYLE_SU_FIVE_LEFT)

	//electricity cost sum
	doc.AddParagraph("")
	costSumPara := doc.AddFormattedParagraph("2.本期用电量：", &styleTextFiveNormal)
	costSumPara.Runs = append(costSumPara.Runs, runWithunderline(fmt.Sprintf(" %.2f ", indic.CostAll)), runNormal("度"))
	//electricity price
	doc.AddParagraph("")
	pricePara := doc.AddFormattedParagraph("3.单价：￥", &styleTextFiveNormal)
	pricePara.Runs = append(pricePara.Runs, runWithunderline(" 1.00 "), runNormal("元/度"))
	//electricity pay
	doc.AddParagraph("")

	para := doc.AddFormattedParagraph("4.本期电费金额：￥ ", &styleTextFiveNormal)
	if companyInfo.RateOfPay != 0 {
		para.Runs = append(para.Runs, runWithunderline(fmt.Sprintf(" %.2f ", indic.CostAll*companyInfo.RateOfPay)), runNormal("元"))
	} else {
		para.Runs = append(para.Runs, runWithunderline(fmt.Sprintf(" %.2f ", indic.CostAll)), runNormal("元"))
	}
	//liquidated damages
	doc.AddParagraph("")
	doc.AddFormattedParagraph("5.违约金（如逾期）：￥__/_____ 元", &styleTextFiveNormal)

	//expense sum
	doc.AddParagraph("")

	expenseSumPara := doc.AddFormattedParagraph("6.合计应缴金额：￥", &styleTextFiveNormal)
	if companyInfo.RateOfPay != 0 {
		expenseSumPara.Runs = append(expenseSumPara.Runs, runWithunderline(fmt.Sprintf(" %.2f ", indic.CostAll*companyInfo.RateOfPay)), runNormal("元"))
	} else {
		expenseSumPara.Runs = append(expenseSumPara.Runs, runWithunderline(fmt.Sprintf(" %.2f ", indic.CostAll)), runNormal("元"))
	}

	//account info
	doc.AddParagraph("")
	doc.AddFormattedParagraph("7.账户信息：", &styleTextFiveNormal)
	doc.AddFormattedParagraph(viper.GetString("account_name"), &styleTextFiveNormal)
	doc.AddFormattedParagraph(fmt.Sprintf("地址：%s", viper.GetString("address")), &styleTextFiveNormal)
	doc.AddFormattedParagraph(fmt.Sprintf("开户行：%s", viper.GetString("account_bank")), &styleTextFiveNormal)
	doc.AddFormattedParagraph(fmt.Sprintf("账号：%s", viper.GetString("account_number")), &styleTextFiveNormal)
}

func runWithunderline(underline string) document.Run {
	return document.Run{
		Text: document.Text{
			Content: underline,
		},
		Properties: &document.RunProperties{
			FontFamily: &document.FontFamily{
				ASCII:    "宋体 (中文正文)",
				HAnsi:    "宋体 (中文正文)",
				EastAsia: "宋体 (中文正文)",
				CS:       "宋体 (中文正文)",
			},
			FontSize: &document.FontSize{
				Val: "22",
			},
			Underline: &document.Underline{
				Val: "single",
			},
		},
	}
}

func runNormal(s string) document.Run {

	return document.Run{
		Text: document.Text{
			Content: s,
		},
		Properties: &document.RunProperties{
			FontFamily: &document.FontFamily{
				ASCII:    "宋体 (中文正文)",
				HAnsi:    "宋体 (中文正文)",
				EastAsia: "宋体 (中文正文)",
				CS:       "宋体 (中文正文)",
			},
			FontSize: &document.FontSize{
				Val: "22",
			},
		},
	}
}

func backup(doc *document.Document) {
	year := viper.GetInt("target_year")
	month := viper.GetInt("target_month")
	next := time.Date(year, time.Month(month), 15, 0, 0, 0, 0, time.UTC)
	next= next.AddDate(0, 1, 0)
	para := doc.AddParagraph(fmt.Sprintf("缴费截止日期\n请在 %d年%d月15日前缴纳，逾期将按日收取违约金（0.05%%-0.1%%/天）。", next.Year(), next.Month()))
	para.SetStyle(STYLE_SU_FIVE_LEFT)
}

func sortByGateNo(companies *map[string]types.CompanyInfo) []string {

	keys := []string{}
	for key := range *companies {
		keys = append(keys, key)
	}
	sort.Strings(keys)

	return keys
}
