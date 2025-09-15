package main

import (
	"context"
	"encoding/json"
	"excel-tools/backend"
	"github.com/wailsapp/wails/v2/pkg/runtime"
	"os"
)

type App struct {
	ctx context.Context

	setting *backend.AppSetting
}

func NewApp() *App {
	return &App{
		setting: &backend.AppSetting{
			KeyWordStatConfig: backend.KeyWordStatConfig{
				KwInputDir:    "",
				KwOutputDir:   "",
				StatMode:      "行模式",
				TargetNumber:  8,
				ForwardNumber: 5,
				SelectedColor: "绿色",
			},
			WordFreqStatConfig: backend.WordFreqStatConfig{
				WfInputDir:     "",
				IntervalNumber: 5,
				SplitChar:      "",
			},
		},
	}
}

func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
	// 加载配置
	a.loadConfig()

}

// 关键字统计相关方法
func (a *App) GetKeyWordStatsConfig() backend.KeyWordStatConfig {
	return a.setting.KeyWordStatConfig
}

func (a *App) UpdateKeyWordStatsConfig(config backend.KeyWordStatConfig) {
	a.setting.KeyWordStatConfig = config
	a.saveConfig()
}

func (a *App) SelectKeyWordInputDir() (string, error) {
	return runtime.OpenDirectoryDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "选择输入文件夹",
	})
}

func (a *App) SelectKeyWordOutputDir() (string, error) {
	return runtime.OpenDirectoryDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "选择输出文件夹",
	})
}

func (a *App) RunKeyWordStats() (string, error) {
	err := a.setting.KeyWordStatConfig.KeyWordStat()
	if err != nil {
		return err.Error(), nil
	}
	return "统计成功", nil
}

// 词频统计相关方法
func (a *App) GetWordFreqStatsConfig() backend.WordFreqStatConfig {
	return a.setting.WordFreqStatConfig
}

func (a *App) UpdateWordFreqStatsConfig(config backend.WordFreqStatConfig) {
	a.setting.WordFreqStatConfig = config
	a.saveConfig()
}

func (a *App) SelectWordFreqInputDir() (string, error) {
	return runtime.OpenDirectoryDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "选择输入文件夹",
	})
}

func (a *App) RunWordFreqStats() (string, error) {
	err := a.setting.WordFreqStatConfig.WordFreqStat()
	if err != nil {
		return err.Error(), nil
	}
	return "统计成功", nil
}

// 配置管理
func (a *App) getConfigPath() string {
	return "config.json"
}

func (a *App) loadConfig() {
	path := a.getConfigPath()
	data, err := os.ReadFile(path)
	if err != nil {
		runtime.LogInfo(a.ctx, "未找到配置文件，使用默认配置")
		return
	}
	err = json.Unmarshal(data, a.setting)
	if err != nil {
		runtime.LogError(a.ctx, "读取配置失败: "+err.Error())
	}
}

func (a *App) saveConfig() {
	path := a.getConfigPath()
	data, err := json.MarshalIndent(a.setting, "", "  ")
	if err != nil {
		runtime.LogError(a.ctx, "保存配置失败: "+err.Error())
		return
	}
	err = os.WriteFile(path, data, 0644)
	if err != nil {
		runtime.LogError(a.ctx, "写入配置文件失败: "+err.Error())
	}
}

func (a *App) ClearMemory() {
	// 重置配置
	a.setting = &backend.AppSetting{
		KeyWordStatConfig: backend.KeyWordStatConfig{
			KwInputDir:    "",
			KwOutputDir:   "",
			StatMode:      "行模式",
			TargetNumber:  8,
			ForwardNumber: 5,
			SelectedColor: "绿色",
		},
		WordFreqStatConfig: backend.WordFreqStatConfig{
			WfInputDir:     "",
			IntervalNumber: 5,
			SplitChar:      "",
		},
	}

	// 保存重置后的配置
	a.saveConfig()
}
