<template>
  <div id="app">
    <header>
      <h1>Excel 工具集</h1>
    </header>

    <main>
      <div class="tabs">
        <button
            :class="{ active: activeTab === 'keyword' }"
            @click="activeTab = 'keyword'"
        >
          关键字统计工具
        </button>
        <button
            :class="{ active: activeTab === 'wordfreq' }"
            @click="activeTab = 'wordfreq'"
        >
          词频统计
        </button>
      </div>

      <div v-if="activeTab === 'keyword'" class="tab-content">
        <h2>关键字统计工具</h2>

        <div class="form-grid">
          <div class="form-group">
            <label>统计模式</label>
            <div class="radio-group">
              <label>
                <input type="radio" v-model="keywordConfig.statMode" value="行统计" @change="updateKeywordConfig" />
                行统计
              </label>
              <label>
                <input type="radio" v-model="keywordConfig.statMode" value="列统计" @change="updateKeywordConfig" />
                列统计
              </label>
            </div>
          </div>

          <div class="form-group">
            <label>目标行/列</label>
            <input type="number" v-model.number="keywordConfig.targetNumber" @change="updateKeywordConfig" />
          </div>

          <div class="form-group">
            <label>向前行/列</label>
            <input type="number" v-model.number="keywordConfig.forwardNumber" @change="updateKeywordConfig" />
          </div>

          <div class="form-group">
            <label>选择颜色</label>
            <select v-model="keywordConfig.selectedColor" @change="updateKeywordConfig">
              <option value="红色">红色</option>
              <option value="绿色">绿色</option>
              <option value="蓝色">蓝色</option>
            </select>
          </div>

          <div class="form-group">
            <label>输入文件夹</label>
            <div class="dir-selector">
              <button @click="selectKeywordInputDir" class="secondary">选择文件夹</button>
              <div class="dir-path">{{ keywordInputDirName }}</div>
            </div>
          </div>

          <div class="form-group">
            <label>输出文件夹</label>
            <div class="dir-selector">
              <button @click="selectKeywordOutputDir" class="secondary">选择输出文件夹</button>
              <div class="dir-path">{{ keywordOutputDirName }}</div>
            </div>
          </div>
        </div>

        <div class="actions">
          <button @click="runKeywordStats" class="primary">生成</button>
          <button @click="clearMemory" class="secondary">清除记忆</button>
        </div>
      </div>

      <div v-if="activeTab === 'wordfreq'" class="tab-content">
        <h2>词频统计工具</h2>

        <div class="form-grid">
          <div class="form-group">
            <label>选择文件夹</label>
            <div class="dir-selector">
              <button @click="selectWordFreqInputDir" class="secondary">选择文件夹</button>
              <div class="dir-path">{{ wordFreqInputDirName }}</div>
            </div>
          </div>

          <div class="form-group">
            <label>分割字符</label>
            <input type="text" v-model="wordFreqConfig.splitChar" @change="updateWordFreqConfig" />
          </div>

          <div class="form-group">
            <label>间隔数量</label>
            <input type="number" v-model.number="wordFreqConfig.intervalNumber" @change="updateWordFreqConfig" />
          </div>
        </div>

        <div class="actions">
          <button @click="runWordFreqStats" class="primary">生成</button>
          <button @click="clearMemory" class="secondary">清除记忆</button>
        </div>
      </div>
    </main>

    <div v-if="message" class="modal">
      <div class="modal-content">
        <h3>结果</h3>
        <p>{{ message }}</p>
        <button @click="message = ''">关闭</button>
      </div>
    </div>
  </div>
</template>

<script>
export default {
  name: 'App',
  data() {
    return {
      activeTab: 'keyword',
      keywordConfig: {
        kwInputDir: '',
        kwOutputDir: '',
        statMode: '行统计',
        targetNumber: 9,
        forwardNumber: 8,
        selectedColor: '绿色'
      },
      wordFreqConfig: {
        wfInputDir: '',
        splitChar: '',
        intervalNumber: 8
      },
      message: ''
    }
  },
  computed: {
    keywordInputDirName() {
      return this.keywordConfig.kwInputDir ? this.keywordConfig.kwInputDir.split(/[\\/]/).pop() : '未选择文件夹';
    },
    keywordOutputDirName() {
      return this.keywordConfig.kwOutputDir ? this.keywordConfig.kwOutputDir.split(/[\\/]/).pop() : '未选择输出文件夹';
    },
    wordFreqInputDirName() {
      return this.wordFreqConfig.wfInputDir ? this.wordFreqConfig.wfInputDir.split(/[\\/]/).pop() : '未选择文件夹';
    }
  },
  async mounted() {
    // 加载配置
    await this.loadConfig();
  },
  methods: {
    async loadConfig() {
      try {
        const kwConfig = await window.go.main.App.GetKeyWordStatsConfig();
        if (kwConfig) {
          this.keywordConfig = kwConfig;
        }

        const wfConfig = await window.go.main.App.GetWordFreqStatsConfig();
        if (wfConfig) {
          this.wordFreqConfig = wfConfig;
        }
      } catch (error) {
        console.error('加载配置失败:', error);
      }
    },

    updateKeywordConfig() {
      window.go.main.App.UpdateKeyWordStatsConfig(this.keywordConfig);
    },

    updateWordFreqConfig() {
      window.go.main.App.UpdateWordFreqStatsConfig(this.wordFreqConfig);
    },

    async selectKeywordInputDir() {
      try {
        const dir = await window.go.main.App.SelectKeyWordInputDir();
        if (dir) {
          this.keywordConfig.kwInputDir = dir;
          this.updateKeywordConfig();
        }
      } catch (error) {
        console.error('选择文件夹失败:', error);
        this.message = "选择文件夹失败: " + error.message;
      }
    },

    async selectKeywordOutputDir() {
      try {
        const dir = await window.go.main.App.SelectKeyWordOutputDir();
        if (dir) {
          this.keywordConfig.kwOutputDir = dir;
          this.updateKeywordConfig();
        }
      } catch (error) {
        console.error('选择输出文件夹失败:', error);
        this.message = "选择输出文件夹失败: " + error.message;
      }
    },

    async selectWordFreqInputDir() {
      try {
        const dir = await window.go.main.App.SelectWordFreqInputDir();
        if (dir) {
          this.wordFreqConfig.wfInputDir = dir;
          this.updateWordFreqConfig();
        }
      } catch (error) {
        console.error('选择文件夹失败:', error);
        this.message = "选择文件夹失败: " + error.message;
      }
    },

    async runKeywordStats() {
      if (!this.keywordConfig.kwInputDir) {
        this.message = "请选择输入文件夹";
        return;
      }

      if (!this.keywordConfig.kwOutputDir) {
        this.message = "请选择输出文件夹";
        return;
      }

      try {
        const result = await window.go.main.App.RunKeyWordStats();
        this.message = result;
      } catch (error) {
        this.message = "执行失败: " + error.message;
      }
    },

    async runWordFreqStats() {
      if (!this.wordFreqConfig.wfInputDir) {
        this.message = "请选择文件夹";
        return;
      }

      if (!this.wordFreqConfig.splitChar) {
        this.message = "请输入分割字符";
        return;
      }

      try {
        const result = await window.go.main.App.RunWordFreqStats();
        this.message = result;
      } catch (error) {
        this.message = "执行失败: " + error.message;
      }
    },

    async clearMemory() {
      try {
        await window.go.main.App.ClearMemory();
        this.keywordConfig = {
          kwInputDir: '',
          kwOutputDir: '',
          statMode: '行统计',
          targetNumber: 9,
          forwardNumber: 8,
          selectedColor: '绿色'
        };
        this.wordFreqConfig = {
          wfInputDir: '',
          splitChar: '',
          intervalNumber: 8
        };
        this.message = "记忆已清除";
      } catch (error) {
        this.message = "清除记忆失败: " + error.message;
      }
    }
  }
}
</script>

<style>
:root {
  --primary-color: #4361ee;
  --secondary-color: #3f37c9;
  --success-color: #4cc9f0;
  --light-color: #f8f9fa;
  --dark-color: #212529;
  --gray-color: #6c757d;
  --border-color: #dee2e6;
  --card-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
  --transition: all 0.3s ease;
}

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
  background-color: #f5f7fb;
  color: var(--dark-color);
  line-height: 1.6;
  padding: 0;
  margin: 0;
}

#app {
  max-width: 1000px;
  margin: 0 auto;
  padding: 20px;
}

header {
  text-align: center;
  margin-bottom: 30px;
  padding: 20px 0;
  background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
  color: white;
  border-radius: 10px;
  box-shadow: var(--card-shadow);
}

h1 {
  font-size: 28px;
  font-weight: 600;
}

h2 {
  font-size: 22px;
  margin-bottom: 20px;
  color: var(--primary-color);
  border-bottom: 2px solid var(--primary-color);
  padding-bottom: 10px;
}

.tabs {
  display: flex;
  background-color: white;
  border-radius: 10px;
  overflow: hidden;
  box-shadow: var(--card-shadow);
  margin-bottom: 30px;
}

.tabs button {
  flex: 1;
  padding: 15px;
  background: none;
  border: none;
  cursor: pointer;
  font-size: 16px;
  font-weight: 500;
  transition: var(--transition);
  color: var(--gray-color);
}

.tabs button.active {
  background-color: var(--primary-color);
  color: white;
}

.tabs button:hover:not(.active) {
  background-color: #f0f0f0;
}

.tab-content {
  background-color: white;
  padding: 30px;
  border-radius: 10px;
  box-shadow: var(--card-shadow);
}

.form-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
  gap: 20px;
  margin-bottom: 25px;
}

.form-group {
  display: flex;
  flex-direction: column;
}

.form-group label {
  font-weight: 600;
  margin-bottom: 8px;
  color: var(--dark-color);
}

.form-group input, .form-group select {
  padding: 12px;
  border: 1px solid var(--border-color);
  border-radius: 6px;
  font-size: 16px;
  transition: var(--transition);
}

.form-group input:focus, .form-group select:focus {
  outline: none;
  border-color: var(--primary-color);
  box-shadow: 0 0 0 3px rgba(67, 97, 238, 0.2);
}

.radio-group {
  display: flex;
  gap: 20px;
}

.radio-group label {
  display: flex;
  align-items: center;
  font-weight: normal;
  cursor: pointer;
}

.radio-group input {
  margin-right: 8px;
}

.dir-selector {
  display: flex;
  align-items: center;
  gap: 10px;
}

.dir-selector button {
  flex-shrink: 0;
}

.dir-path {
  padding: 10px;
  background-color: #f8f9fa;
  border-radius: 6px;
  border: 1px dashed var(--border-color);
  flex-grow: 1;
  min-height: 44px;
  display: flex;
  align-items: center;
}

.actions {
  display: flex;
  justify-content: flex-end;
  gap: 15px;
  margin-top: 20px;
  padding-top: 20px;
  border-top: 1px solid var(--border-color);
}

button {
  padding: 12px 24px;
  border: none;
  border-radius: 6px;
  cursor: pointer;
  font-size: 16px;
  font-weight: 500;
  transition: var(--transition);
}

button.primary {
  background-color: var(--primary-color);
  color: white;
}

button.primary:hover {
  background-color: var(--secondary-color);
}

button.secondary {
  background-color: #e9ecef;
  color: var(--dark-color);
}

button.secondary:hover {
  background-color: #dee2e6;
}

.modal {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background-color: rgba(0, 0, 0, 0.5);
  display: flex;
  justify-content: center;
  align-items: center;
  z-index: 1000;
}

.modal-content {
  background-color: white;
  padding: 30px;
  border-radius: 10px;
  max-width: 500px;
  width: 90%;
  box-shadow: 0 10px 25px rgba(0, 0, 0, 0.2);
}

.modal-content h3 {
  margin-bottom: 15px;
  color: var(--primary-color);
}

.modal-content p {
  white-space: pre-line;
  margin-bottom: 20px;
  line-height: 1.5;
}

.modal-content button {
  width: 100%;
  background-color: var(--primary-color);
  color: white;
}

@media (max-width: 768px) {
  #app {
    padding: 15px;
  }

  .form-grid {
    grid-template-columns: 1fr;
  }

  .tabs {
    flex-direction: column;
  }
}
</style>