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
	STYLE_TABLE_MULTI_ROOMS
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

	styleTableMultiRows = document.TableConfig{
		Cols:      5,
		Width:     9000,
		ColWidths: []int{1800, 1800, 1800, 1800, 1800},
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

func handleDocxCreate(unitFinish *chan string, count int, unitWG *sync.WaitGroup) {
	unitCount := count

	for {
		str := <-*unitFinish
		log.Println("recive msg", str)
		if strings.HasPrefix(str, types.FINISH_FLAG) {
			log.Println("------------done--------", str)
			unitWG.Done()
			unitCount--
		} else {
			log.Println("recive none finish msg ", str)
		}

		if unitCount == 0 {
			return
		}
	}
}

func CreateDocxs(unitToNotiMap *map[int][]types.NotificationItem, finish *chan string) {
	document.SetGlobalLevel(document.LogLevelError)
	path := viper.GetString("output")

	unitFinishChan := make(chan string)

	filepath.Walk(path, func(path string, info fs.FileInfo, err error) error {
		_, e := os.Stat(path)
		if os.IsNotExist(e) {
			return nil
		}
		fileHead := fmt.Sprintf("%d-%d电费通知单", viper.GetInt("target_year"), viper.GetInt("target_month"))
		if info != nil && strings.HasPrefix(info.Name(), fileHead) && strings.HasSuffix(info.Name(), ".docx") {
			log.Println("remove ", path)
			os.Remove(path)
		}
		return nil
	})

	unitCount := len(*unitToNotiMap)
	var unitWG sync.WaitGroup
	unitWG.Add(unitCount)
	go handleDocxCreate(&unitFinishChan, unitCount, &unitWG)

	for unit, items := range *unitToNotiMap {
		go createBillsByUnit(&items, unit, &unitFinishChan)
	}
	unitWG.Wait()
	*finish <- fmt.Sprintf("docx_create_%s", types.FINISH_FLAG)
}

func createBillsByUnit(notis *[]types.NotificationItem, unit int, finish *chan string) {
	if len(*notis) == 0 {
		*finish <- fmt.Sprintf("%s_unit_%d_empty", types.FINISH_FLAG, unit)
		return
	}
	doc := document.New()
	setUpStyle(doc)
	doc.SetPageSize(document.PageSizeA4)
	doc.SetPageOrientation(document.OrientationPortrait)
	doc.SetPageMargins(23, 27, 23, 27)

	// utils.SortByGateNo(notis)

	for i, noti := range *notis {
		docxPage(doc, &noti)
		if i != len(*notis)-1 {
			doc.AddPageBreak()
		}
	}

	err := doc.Save(path.Join(viper.GetString("output"), fmt.Sprintf("%d-%d电费通知单-%s单元.docx", viper.GetInt("target_year"), viper.GetInt("target_month"), CN_NUMBER[unit])))
	fmt.Println(err)
	*finish <- fmt.Sprintf("%s_docx_unit_%d", types.FINISH_FLAG, unit)

}

func docxPage(doc *document.Document, noti *types.NotificationItem) {
	title(doc)
	doc.AddParagraph("")
	floor(doc, &noti.CompanyInfo)
	nameAndAddress(doc, &noti.CompanyInfo)
	doc.AddParagraph("")
	doc.AddParagraph("")
	sign(doc)
	doc.AddParagraph("")
	billTable(doc, noti)
	expense2(doc, noti)
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

func tableArea2(doc *document.Document, noti *types.NotificationItem, style int) {
	var config document.TableConfig = document.TableConfig{}
	tableTitle(doc, &noti.CompanyInfo, style)
	tableConfig2(noti, &config, style)
	if len(config.Data) == 0 {
		log.Fatal("unsupport table style ", style)
		return
	}
	table, _ := doc.AddTable(&config)

	for row := 0; row < table.GetRowCount(); row++ {
		if config.Rows > 3 {
			table.SetRowHeight(row, &document.RowHeightConfig{
				Height: 19,
			})
		} else {
			table.SetRowHeight(row, &document.RowHeightConfig{
				Height: 26,
			})
		}

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
	case STYLE_TABLE_MULTI_ROOMS:
		paraTitle = doc.AddParagraph(fmt.Sprintf("%d年%d月用电量", viper.GetInt("target_year"), viper.GetInt("target_month")))
	}
	paraTitle.SetStyle(STYLE_SU_THREE_CENTER_B)

}

func tableConfig2(noti *types.NotificationItem, config *document.TableConfig, style int) {
	switch style {
	case STYLE_TABLE_BILL:
		config.Cols = styleTableBillNormal.Cols
		config.Rows = styleTableBillNormal.Rows
		config.Width = styleTableBillNormal.Width
		config.ColWidths = styleTableBillNormal.ColWidths
		config.Data = [][]string{
			{"月份", "上月表数", "倍率", "本月表数", "实际用量（度）", "应缴电费"},
			{fmt.Sprint(viper.GetInt("target_month")),
				fmt.Sprint((*noti.IndicList)[0].IndicLastMonth),
				strconv.FormatFloat((*noti.IndicList)[0].Times, 'f', 0, 64),
				strconv.FormatFloat((*noti.IndicList)[0].IndicCurrent, 'f', 2, 64),
				strconv.FormatFloat((*noti.IndicList)[0].Cost, 'f', 2, 64),
				strconv.FormatFloat((*noti.IndicList)[0].Payment, 'f', 2, 64)},
		}

	case STYLE_TABLE_BILL_NO_PAYMENT:
		config.Cols = styleTableBillNoPay.Cols
		config.Rows = styleTableBillNoPay.Rows
		config.Width = styleTableBillNoPay.Width
		config.ColWidths = styleTableBillNoPay.ColWidths
		config.Data = [][]string{
			{"月份", "上月表数", "倍率", "本月表数", "实际用量（度）"},
			{fmt.Sprint(viper.GetInt("target_month")),
				fmt.Sprint((*noti.IndicList)[0].IndicLastMonth),
				strconv.FormatFloat((*noti.IndicList)[0].Times, 'f', 0, 64),
				strconv.FormatFloat((*noti.IndicList)[0].IndicCurrent, 'f', 2, 64),
				strconv.FormatFloat((*noti.IndicList)[0].Cost, 'f', 2, 64)},
		}
	case STYLE_TABLE_AIR_CONTROL:
		config.Cols = styleTableBillAirControl.Cols
		config.Rows = styleTableBillAirControl.Rows
		config.Width = styleTableBillAirControl.Width
		config.ColWidths = styleTableBillAirControl.ColWidths
		config.Data = [][]string{
			{"月份", "实际用量（度）", "应缴电费"},
			{fmt.Sprint(viper.GetInt("target_month")),
				strconv.FormatFloat((*noti.AirIndicList)[0].Cost, 'f', 2, 64),
				strconv.FormatFloat((*noti.AirIndicList)[0].Payment, 'f', 2, 64)},
		}
	case STYLE_TABLE_AIR_CONTROL_NO_PAYMENT:
		config.Cols = styleTableBillAirControlNoBill.Cols
		config.Rows = styleTableBillAirControlNoBill.Rows
		config.Width = styleTableBillAirControlNoBill.Width
		config.ColWidths = styleTableBillAirControlNoBill.ColWidths
		config.Data = [][]string{
			{"月份", "实际用量（度）"},
			{fmt.Sprint(viper.GetInt("target_month")), strconv.FormatFloat((*noti.AirIndicList)[0].Cost, 'f', 2, 64)},
		}
	case STYLE_TABLE_MULTI_ROOMS:
		config.Cols = styleTableMultiRows.Cols
		config.Rows = len(noti.GateNos) + 1 //列表长度 + 表头
		config.Width = styleTableMultiRows.Width
		config.ColWidths = styleTableMultiRows.ColWidths
		data := [][]string{{"门牌号", "上月表数", "倍率", "本月表数", "实际用量（度）"}}
		for _, tableRow := range *noti.IndicList {
			row := []string{tableRow.RoomNo,
				strconv.FormatFloat(tableRow.IndicLastMonth, 'f', 2, 64),
				strconv.FormatFloat(tableRow.Times, 'f', 2, 64),
				strconv.FormatFloat(tableRow.IndicCurrent, 'f', 2, 64),
				strconv.FormatFloat(tableRow.Cost, 'f', 2, 64)}
			data = append(data, row)
		}
		config.Data = data
	default:
		//ignore
	}
}

func billTable(doc *document.Document, noti *types.NotificationItem) {
	para := doc.AddParagraph("缴费信息")
	para.SetStyle(STYLE_SU_FIVE_LEFT)

	doc.AddParagraph("")

	if len(noti.GateNos) > 1 {
		tableArea2(doc, noti, STYLE_TABLE_MULTI_ROOMS)
		doc.AddParagraph("")
		return
	}

	if noti.IsAddPayment {
		tableArea2(doc, noti, STYLE_TABLE_BILL)
		doc.AddParagraph("")
	} else {
		tableArea2(doc, noti, STYLE_TABLE_BILL_NO_PAYMENT)
		doc.AddParagraph("")
	}

	if len(*noti.AirIndicList) == 0 {
		return
	}

	if noti.IsAddPayment {
		tableArea2(doc, noti, STYLE_TABLE_AIR_CONTROL)
		doc.AddParagraph("")
	} else {
		tableArea2(doc, noti, STYLE_TABLE_AIR_CONTROL_NO_PAYMENT)
		doc.AddParagraph("")
	}
}

func expense2(doc *document.Document, noti *types.NotificationItem) {
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
	costSumPara.Runs = append(costSumPara.Runs, runWithUnderline(fmt.Sprintf(" %.2f ", noti.CostSum)), runNormal("度"))
	//electricity price
	doc.AddParagraph("")
	pricePara := doc.AddFormattedParagraph("3.单价：￥", &styleTextFiveNormal)
	pricePara.Runs = append(pricePara.Runs, runWithUnderline(" 1.00 "), runNormal("元/度"))
	//electricity pay
	doc.AddParagraph("")

	para := doc.AddFormattedParagraph("4.本期电费金额：￥ ", &styleTextFiveNormal)
	if noti.RateOfPay != 0 {
		para.Runs = append(para.Runs, runWithUnderline(fmt.Sprintf(" %.2f ", noti.CostSum*noti.RateOfPay)), runNormal("元"))
	} else {
		para.Runs = append(para.Runs, runWithUnderline(fmt.Sprintf(" %.2f ", noti.CostSum)), runNormal("元"))
	}
	//liquidated damages
	doc.AddParagraph("")
	doc.AddFormattedParagraph("5.违约金（如逾期）：￥__/_____ 元", &styleTextFiveNormal)

	//expense sum
	doc.AddParagraph("")

	expenseSumPara := doc.AddFormattedParagraph("6.合计应缴金额：￥", &styleTextFiveNormal)
	if noti.RateOfPay != 0 {
		expenseSumPara.Runs = append(expenseSumPara.Runs, runWithUnderline(fmt.Sprintf(" %.2f ", noti.CostSum*noti.RateOfPay)), runNormal("元"))
	} else {
		expenseSumPara.Runs = append(expenseSumPara.Runs, runWithUnderline(fmt.Sprintf(" %.2f ", noti.CostSum)), runNormal("元"))
	}

	//account info
	doc.AddParagraph("")
	doc.AddFormattedParagraph("7.账户信息：", &styleTextFiveNormal)
	doc.AddFormattedParagraph(viper.GetString("account_name"), &styleTextFiveNormal)
	doc.AddFormattedParagraph(fmt.Sprintf("地址：%s", viper.GetString("address")), &styleTextFiveNormal)
	doc.AddFormattedParagraph(fmt.Sprintf("开户行：%s", viper.GetString("account_bank")), &styleTextFiveNormal)
	doc.AddFormattedParagraph(fmt.Sprintf("账号：%s", viper.GetString("account_number")), &styleTextFiveNormal)
}

func runWithUnderline(underline string) document.Run {
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
	next = next.AddDate(0, 1, 0)
	para := doc.AddParagraph(fmt.Sprintf("缴费截止日期\n请在 %d年%d月15日前缴纳，逾期将按日收取违约金（0.05%%-0.1%%/天）。", next.Year(), next.Month()))
	para.SetStyle(STYLE_SU_FIVE_LEFT)
}
