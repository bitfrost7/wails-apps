package backend

import (
	"sort"
	"strings"
)

// 将字符串分成数组
func singleCh(s string) []string {
	return strings.Split(s, "")
}

// 词频统计
func wordFreqStat(words []string, less func(string, string) bool) []*WordFreq {
	freqM := make(map[string]int)
	for _, word := range words {
		freqM[word]++
	}
	res := make([]*WordFreq, 0)
	for word, freq := range freqM {
		res = append(res, &WordFreq{word, freq})
	}
	if less == nil {
		return res
	}
	sort.Slice(res, func(i, j int) bool {
		return less(res[i].word, res[j].word)
	})
	return res
}

func FreqSum(freq []*WordFreq) int {
	sum := 0
	for _, f := range freq {
		sum += f.freq
	}
	return sum
}
