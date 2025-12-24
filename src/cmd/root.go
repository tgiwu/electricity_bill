package cmd

import (
	"bytes"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"sync"

	"electricity_bill/src/business"
	"electricity_bill/src/conf"
	"electricity_bill/src/types"
	"electricity_bill/src/word"

	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

var (
	//配置文件
	cfgFile string
	//输出地址
	output string
	//输入文件
	input string
	//并发线程数
	count int = 0
	//create docx flag
	isCreateDocx = false

	indicMap     map[int]map[string]types.Indication
	companiesMap map[int]map[string]types.CompanyInfo
	rootCmd      = &cobra.Command{
		Use:   "electricity bill",
		Short: "construct electricity bill file",
		Long:  "construct electricity bill file",
		Run: func(cmd *cobra.Command, args []string) {

			fmt.Printf("all settings : %+v\n", viper.AllSettings())

			fmt.Println("input  \n", viper.GetString("input"))

			err := checkCriticalParameter()

			if err != nil {
				log.Panicln(err)
			}
			cc := make(chan types.CompanyInfo)
			ce := make(chan types.Indication)
			cFinish := make(chan (string))

			//read company info
			count++
			//read indication info
			count++
			var wg sync.WaitGroup
			wg.Add(count)

			go handleChan(cc, cFinish, ce, &wg)
			go business.ReadCompany(&cc, &cFinish)
			go business.ReadElec(&ce, &cFinish)

			wg.Wait()

			if len(companiesMap) > 0 && len(indicMap) > 0 {

				wg = sync.WaitGroup{}
				count = len(companiesMap)
				wg.Add(len(companiesMap))

				go handleDocxCreate(cFinish, &wg)

				word.CreateDocxs(&indicMap, &companiesMap, &cFinish, &wg)
				wg.Wait()
			}

		},
	}
)

func checkCriticalParameter() error {
	outputDir := viper.GetString("output")
	_, err := os.Stat(outputDir)
	if err != nil && os.IsNotExist(err) {
		log.Println("build output folder")
		err = os.MkdirAll(outputDir, os.FileMode(0777))
		if err != nil {
			return err
		}
	}

	//indication file
	indicFilePath := viper.GetString("elec_file")

	_, err = os.Stat(indicFilePath)
	if err != nil {
		switch {
		case os.IsNotExist(err):
			return types.MyError{
				Path: indicFilePath,
				Op:   "indic file not exist",
			}
		case !os.IsPermission(err):
			return types.MyError{
				Path: indicFilePath,
				Op:   "need more permission",
			}
		}
	}

	if !strings.HasSuffix(indicFilePath, ".xlsx") {
		return types.MyError{
			Path: indicFilePath,
			Op:   "err file type, not a xlsx file",
		}
	}
	return nil
}

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
	}
}

func init() {
	cobra.OnInitialize(initConfig)
	currentPath, _ := filepath.Abs(".")

	viper.SetDefault("elec_file", filepath.Join(currentPath, "electricity.xlsx"))
	viper.SetDefault("output", currentPath)
	viper.SetDefault("target_month", 11)
	viper.SetDefault("target_year", "2025")
	viper.SetDefault("company_sheet", "公司信息")
	viper.SetDefault("indication_sheet", "电量统计")

	rootCmd.PersistentFlags().StringVarP(&cfgFile, "config", "c", "", "config file")
	rootCmd.PersistentFlags().StringVarP(&output, "output", "o", currentPath, "output path")
	rootCmd.PersistentFlags().StringVarP(&input, "elec_file", "i", filepath.Join(currentPath, "electricity.xlsx"), "input file")
	rootCmd.PersistentFlags().IntP("target_month", "m", 1, "target month")
	rootCmd.PersistentFlags().StringP("target_year", "y", "2025", "target year")

	viper.BindPFlag("output", rootCmd.PersistentFlags().Lookup("output"))
	viper.BindPFlag("elec_file", rootCmd.PersistentFlags().Lookup("elec_file"))
	viper.BindPFlag("target_month", rootCmd.PersistentFlags().Lookup("target_month"))
	viper.BindPFlag("target_year", rootCmd.PersistentFlags().Lookup("target_year"))

}

func initConfig() {

	conf.ReadConfig()

	if cfgFile != "" {

		bs, err := os.ReadFile(cfgFile)

		viper.MergeConfig(bytes.NewReader(bs))

		if err != nil {
			panic(err)
		}
	}

	fmt.Printf("init config  %+v \n", viper.AllSettings())
}

func handleDocxCreate(cFinish chan (string), wg *sync.WaitGroup) {
	for {
		str := <-cFinish
		log.Println("recive msg", str)
		if strings.HasPrefix(str, "_f") {
			log.Println("------------done--------", str)
			wg.Done()
			count--
		} else {
			log.Println("recive none finish msg ", str)
		}

		if count == 0 {
			return
		}
	}
}

func handleChan(cc chan (types.CompanyInfo), cFinish chan (string), ce chan types.Indication, wg *sync.WaitGroup) {
	for {
		select {
		case info := <-cc:
			// fmt.Printf("%+v \n", info)
			if len(info.GateNo) == 0 {
				log.Printf("illegal room no %+v \n", info)
				continue
			}
			// unit := strings.Split(info.GateNo, "-")[0]
			if companiesMap == nil {
				companiesMap = make(map[int]map[string]types.CompanyInfo)
			}

			if unitCompaniesMap, found := companiesMap[info.Unit]; !found {
				unitCompaniesMap = make(map[string]types.CompanyInfo)
				unitCompaniesMap[info.GateNo] = info
				companiesMap[info.Unit] = unitCompaniesMap
			} else {
				unitCompaniesMap[info.GateNo] = info
				companiesMap[info.Unit] = unitCompaniesMap
			}

		case str := <-cFinish:
			fmt.Println(str)
			if str == "ele_f" || str == "com_f" || str == "doc_f" {
				wg.Done()
				count--
			}

			if count == 0 {
				log.Println("read finish , start create docx")
				return
			}

		case indic := <-ce:
			// fmt.Printf("%+v\n", indic)
			if len(indic.RoomNo) == 0 {
				log.Printf("illegal room no %+v \n", indic)
				continue
			}

			if indicMap == nil {
				indicMap = make(map[int]map[string]types.Indication, 1)
				unitIndicMap := make(map[string]types.Indication, 1)
				unitIndicMap[indic.RoomNo] = indic
				indicMap[indic.Unit] = unitIndicMap
				continue
			}

			unitIndicMap, found := indicMap[indic.Unit]

			if !found {
				unitIndicMap = make(map[string]types.Indication, 0)
			}

			unitIndicMap[indic.RoomNo] = indic
			indicMap[indic.Unit] = unitIndicMap
		}
	}
}
