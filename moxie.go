package main

import (
	"fmt"
	"os"
	"path/filepath"

	"word/tool"
)

func main() {
	// 创建excel目录
	excelDir := "excel"
	if _, err := os.Stat(excelDir); os.IsNotExist(err) {
		err := os.Mkdir(excelDir, 0755)
		if err != nil {
			fmt.Printf("创建目录失败: %v\n", err)
			return
		}
	}

	// 合并所有单词到一个列表中
	allWords := []string{
		"n.祖父",
		"n.祖母",
		"n.父亲",
		"n.母亲",
		"n.伯父，叔叔",
		"n.伯母，婶婶",
		"n.表兄弟姐妹",
		"n.男人",
		"n.女人",
		"n.卧室",
		"n.房子",
		"n.厨房",
		"n.起居室，客厅",
		"n.浴室",
		"prep.在…之间",
		"n.餐厅",
		"prep.在…旁边",
		"n.床",
		"n.椅子",
		"n.计算机，电脑",
		"n.书桌",
		"n.台灯",
		"n.海报",
		"n.架子",
		"n.门",
		"n.墙",
		"n.早晨",
		"n.下午",
		"n.傍晚",
		"n.晚上",
		"n.早餐",
		"n.午餐",
		"n.晚饭",
		"num.十一",
		"num.十二",
		"n.船",
		"n.游戏用具",
		"n.手偶",
		"n.拼图游戏",
		"n.滑板",
		"n.跳绳",
		"n.宇宙飞船",
		"n.足球",
		"n.飞盘",
		"n.鸟",
		"n.鸡",
		"n.驴",
		"n.鸭子",
		"n.农场",
		"n.山 goat",
		"n.马",
		"n.羊羔",
		"n.绵羊",
		"n.公牛",
		"n.奶牛",
		"n.盒子",
		"n.樱桃",
		"n.盘子",
		"n.食物",
		"n.叉子",
		"n.玻璃杯",
		"n.刀",
		"n.三明治",
		"n.草莓",
		"n.竹子",
		"n.桃子",
		"n.土豆",
		"n.连衣裙",
		"n.夹克",
		"n.牛仔裤",
		"n.鞋",
		"n.半身裙",
		"n.短袜",
		"n.裤子",
		"n.外套",
		"n.围巾",
		"adj.干净的",
		"adj.脏的",
		"adj.干的",
		"adj.湿的",
		"adj.高兴的",
		"adj.不高兴的",
		"adj.饿的",
		"adj.渴的",
		"adj.冷的",
		"adj.热的",
		"adj.新的",
		"adj.旧的",
		"n.乡下",
		"n.田地",
		"n.花",
		"n.山丘",
		"n.河",
		"n.城镇",
		"n.树",
		"n.公共汽车",
		"n.路",
		"n.商店",
		"num.十三",
		"num.十四",
		"num.十五",
		"num.十六",
		"num.十七",
		"num.十八",
		"num.十九",
		"num.二十",
		"num.二十一",
		"n.饼干",
		"n.灌木丛",
		"v.接住",
		"v.喝",
		"v.吃",
		"v.藏",
		"v.跳",
		"v.跑",
		"坐下",
		"站起来",
		"n.水",
		"v.喊",
		"v.唱歌",
		"n.歌曲",
		"v.扔",
		"v.画画",
		"n.跳舞",
		"vi.听",
		"踢足球",
		"弹钢琴",
		"n.网球",
		"拉小提琴",
		"v.读",
		"v.写",
		"n.手",
		"n.脚趾",
		"v.触碰",
	}

	// 定义数据集
	allWordsList := allWords

	filename := filepath.Join(excelDir, "新概念-青少版B.xlsx")

	if err := tool.GenExerSheet(allWordsList, filename); err != nil {
		fmt.Printf("生成Excel文件时发生错误: %v\n", err)
	} else {
		fmt.Printf("成功生成Excel文件: %s\n", filename)
	}
}
