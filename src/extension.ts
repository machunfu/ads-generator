// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import path = require('path');
import * as vscode from 'vscode';
import * as fs from 'fs';
const XLSX = require('xlsx');

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	// Use the console to output diagnostic information (console.log) and errors (console.error)
	// This line of code will only be executed once when your extension is activated
	console.log('Congratulations, your extension "ads" is now active!');

	// The command has been defined in the package.json file
	// Now provide the implementation of the command with registerCommand
	// The commandId parameter must match the command field in package.json

	context.subscriptions.push(vscode.commands.registerCommand('ads.excel', () => {
		readExcel()
	}))

}


function readExcel() {
	// 打开弹框
	vscode.window.showOpenDialog({
		canSelectFiles: true,
		canSelectFolders: false,
		canSelectMany: false,
		filters: {
			files: ['xlsx']
		},
		title: '请选择ADS配置表(excel)'
	})
		.then(result => {
			if (result === undefined) {
				vscode.window.showInformationMessage('未选择文件');
			} else {

				const filePath = result[0].fsPath
				// 读取Excel文件
				const workbook = XLSX.readFile(filePath);
				// 获取工作表列表
				const sheet_name_list = workbook.SheetNames;
				// 根据工作表名称获取第一个工作表
				const worksheet = workbook.Sheets[sheet_name_list[0]];
				// 将工作表数据转换为JSON对象数组
				const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

				// 创建输出目录
				const directoryPath = path.dirname(filePath)
				const outPutPath = path.join(directoryPath, "vscode_output_files");
				fs.mkdirSync(outPutPath, { recursive: true });

				// 查找自定义组件,并生成自定义组件需要的文件名称
				let list: any[] = []
				data.forEach((row: any[]) => {
					row.forEach(cell => {
						if (cell.includes('mv-components')) {
							// 示例：mv-components-set-campaign-num@配置计划数量@setCampaignNum
							const arr = cell.split('@')

							// 转换fileName：mv-components-set-campaign-num=>MvComponentsSetCampaignNum
							let fileName = ''
							arr[0].split('-').forEach((item: string) => {
								const str = item.charAt(0).toUpperCase() + item.slice(1);
								fileName += str
							})

							// 转换customFolderName：mv-components-set-campaign-num=>set-campaign-num
							let customFolderName = arr[0].replace('mv-components-', '')

							// 转换customVueName：mv-components-set-campaign-num=>SetCampaignNum
							let customVueName = fileName.replace('MvComponents', '')

							// 中文名称
							let cname = arr[1]


							list.push({ fileName, customFolderName, customVueName, cname })
						}
					})
				});

				// 输出壳子文件
				list.forEach(file => {
					const filePath = path.join(outPutPath, `${file.fileName}.vue`)
					// 创建文件并写入内容
					fs.writeFile(filePath, outputMvFileContent(file), (err) => {
						if (err) throw err;
						vscode.window.showInformationMessage('壳子文件 error');
					});
				})

				// 输出自定义组件
				list.forEach(file => {
					const outPutComponentsPath = path.join(outPutPath, "components", file.customFolderName);
					fs.mkdirSync(outPutComponentsPath, { recursive: true });
					const filePath = path.join(outPutComponentsPath, 'index.vue')
					// 创建文件并写入内容
					fs.writeFile(filePath, outputCustomVueFileContent(file), (err) => {
						if (err) throw err;
						vscode.window.showInformationMessage('自定义组件 error');
					});
				})

				// 输入readme文件
				let readmeContent = ''
				list.forEach(file => {
					readmeContent += `${file.fileName} ${file.cname} \r\n`
				})
				// 创建文件并写入内容
				fs.writeFile(path.join(outPutPath, 'README.txt'), readmeContent, (err) => {
					if (err) throw err;
					vscode.window.showInformationMessage('readme error');
				});

				vscode.window.showInformationMessage(`${list.length}个自定义组件生成完毕，输出目录：${outPutPath}`);
			}
		})
}

function outputMvFileContent(file: { fileName: any; customFolderName: any; customVueName: any; }) {
	const { fileName, customFolderName, customVueName } = file
	let str = `<script>
	import eleMixin from '../eleMixin'
	import { setFormData } from '../utils/utils'
	import ${customVueName} from './components/${customFolderName}'
	export default {
	  name: '${fileName}',
	  components: { ${customVueName} },
	  mixins: [eleMixin],
	  methods: {
		handelChange(data) {
		  setFormData(this, data)
		}
	  },
	  render() {
		// These properties come from @eleMixin
		let { eleHidden, initData, initMapData, initList, formData, required, title, name } = this
		return eleHidden ? null : (
		  <${customFolderName}
			v-model={formData}
			init-data={initData}
			init-map-data={initMapData}
			init-list-data={initList}
			title={title}
			name={name}
			required={required}
			onOn-change={this.handelChange}
			onOn-validate={this.handleValidate}></${customFolderName}>
		)
	  }
	}
	</script>
	`

	return str
}

function outputCustomVueFileContent(file: { customFolderName: any; customVueName: any; }) {
	const { customFolderName, customVueName } = file
	let str = `
	<template>
	<div>${customFolderName}</div>
	</template>
	<script>
	import { mapGetters } from 'vuex'
	import componentsMixin from '@/m-form/src/componentsMixin'
	export default {
	  name: '${customVueName}',
	  mixins: [componentsMixin],
	  props: {
		title: {
		  type: String,
		  default: ''
		},
		required: {
		  type: Boolean,
		  default: false
		},
		value: {
		  type: Object,
		  default: () => {}
		},
		name: {
		  type: String,
		  default: ''
		}
	  },
	  data() {
		return {}
	  },
	  computed: {
		...mapGetters(['bluePrint', 'bluePrintService', 'mediaAccountList', 'promotedObjectType'])
	  },
	  methods: {}
	}
	</script>
	`
	return str
}



// This method is called when your extension is deactivated
export function deactivate() { }
