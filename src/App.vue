<template>
    <div id="app">
        <h2>餐补报销表单自动生成工具</h2>
        <div class="handle-box">
            <input
                    type="file"
                    class="upload"
                    @change="handleChange"
            />
            <el-button
                    type="primary"
                    class="upload-button"
            >
                上传考勤文件
            </el-button>
            <div>
                需准备发票: <span>{{money}}</span>元
            </div>
        </div>
        <div class="intro-box">
            <el-card>
                <div slot="header">
                  <span>
                    用户指南
                  </span>
                </div>
                <div class="intro">
                    <p>1 请先在PS系统中导出一段时间考勤excel表</p>
                    <p>2 上传文件自动生成报销表单</p>
                </div>

            </el-card>
        </div>
    </div>
</template>

<script>
    import XLSX from 'xlsx'

    export default {
        name: 'App',
        data() {
            return {
                file: null,
                money: 0
            }
        },
        methods: {
            handleChange(e) {
                let file;
                // 当重选excel且excel存在，更新file
                if (e && e.target.files[0]) {
                    this.file = e.target.files[0]
                }
                file = this.file;
                if (file) {
                    let reader = new FileReader();
                    reader.readAsBinaryString(file);
                    let that = this;
                    reader.onload = function (ev) {
                        let data = ev.target.result;
                        let workbook = XLSX.read(data, {
                            type: 'binary',
                            cellDates: true,
                            dateNF: 'yyyy/mm/dd hh:mm:ss'
                        });
                        let persons = [];
                        persons = persons.concat(XLSX.utils.sheet_to_json(workbook.Sheets.Sheet1));
                        let timeArr = [];
                        let current = null;
                        persons.map(e => {
                            let time = e['打卡时间'];
                            // identify the format of file
                            if (typeof (time) === "string") {
                                alert("导出格式异常")
                            } else {
                                time = new Date(time);
                                let year = time.getFullYear();
                                let month = time.getMonth() + 1;
                                let day = time.getDate();
                                let hour = time.getHours();
                                let minute = time.getMinutes();
                                time = year + '-' + month + '-' + day + '-' + hour + ':' + minute;
                                current = Object.assign({date: month + '-' + day, hour, time}, e)
                            }
                            // put the time which in one day into one array
                            if (timeArr.length === 0) {
                                timeArr.push([current])
                            } else {
                                if (timeArr[timeArr.length - 1][0].date === current.date) {
                                    let currentDay = timeArr[timeArr.length - 1];
                                    currentDay.push(current);
                                    timeArr[timeArr.length - 1] = currentDay
                                } else if (current.hour > 4) {
                                    timeArr.push([current])
                                }
                            }
                        });
                        let result = [];
                        timeArr.map(currentDay => {
                            let currentDayLast = currentDay[currentDay.length - 1];
                            if (currentDayLast['进/出'] == '出') {
                                if (currentDayLast.hour >= 20) {
                                    result.push(currentDayLast)
                                }
                            } else {
                                alert("考勤数据异常")
                            }
                        });

                        result.map(e => {
                            e.domain = e["部门名称"].split(`"`)[1].split("团队")[1];
                            e.group = "无";
                            e.name = e["姓名"].split(`"`)[1];
                            e.showTime = e.time;
                            let redundantAttributes = ['员工 ID','打卡时间','终端描述','终端编码','部门名称','记录产生方式','姓名','进/出','date', 'hour', 'time'];
                            redundantAttributes.map(att => {
                                delete e[att]
                            })
                        });
                        that.money = result.length * 30;
                        let str = `领域，分组，姓名，打卡时间，进\\出\n`;
                        // 增加\t为了不让表格显示科学计数法或者其他格式
                        for (let i = 0; i< result.length; i++) {
                            for (let item in result[i]) {
                                str += `${result[i][item] + '\t'},`;
                            }
                            str += '\n';
                        }
                        // 解决中文乱码
                        let uri = 'data:text/csv;charset=utf-8,\ufeff' + encodeURIComponent(str);
                        let link = document.createElement("a");
                        link.href = uri;
                        link.download = "result.csv";
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);
                    }
                }
            }
        }
    }
</script>

<style>
    #app {
        font-family: Avenir, Helvetica, Arial, sans-serif;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
        text-align: center;
        color: #2c3e50;
        margin-top: 60px;
    }

    .handle-box {
        margin: 100px 0;
    }

    .intro {
        text-align: left;
    }

    .intro-box {
        width: 400px;
        margin: 0 auto;
    }

    .result p span {
        color: darkorange;
        margin: 0 10px;
        font-weight: bold;
    }

    .upload-button {
        position: relative;
        margin-left: -255px !important;
        z-index: -1;
        cursor: pointer;
    }

    input[type="file"] {
        line-height: 80px;
        text-align: center !important;
        opacity: 0;
        cursor: pointer;
    }


</style>
