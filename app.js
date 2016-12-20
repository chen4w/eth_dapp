const async = require('async');
const XLSX = require('xlsx');
const github = require('octonode');

async.waterfall([
	function(callback){
        //get github rep url from the xlsx file
        let workbook = XLSX.readFile('eth_dapp.xlsx');
        let sheetNames = workbook.SheetNames; 
        let worksheet = workbook.Sheets[sheetNames[0]];
        let git_urls = [];
        for (z in worksheet) {
            /* 带!的属性（比如!ref）是表格的特殊属性，用来输出表格的信息，不是表格的内容，所以去掉 */
            if(z[0] !='D') continue;
            let zv = worksheet[z].v;
            if(zv.indexOf('https://github.com/')!=0)
                continue;
            git_urls.push(zv);
        }       
		callback(null, git_urls, 'two');
	},
	function(git_urls, arg2, callback){
	  //get github summery by urls
      //git_urls = git_urls.splice(0,20);
        console.log(git_urls.length+' git repos');
        var client = github.client('10b6124411eec446990b63ef43af7008e5317a2b');
        let rmap = {};
        let cout_fail=0;
        let cout_succ=0;
        //const asy = require('async');
        async.mapLimit(git_urls, 2, function(url, cb) {
            let gurl = url.substring(19);
            var ghrepo = client.repo(gurl);
            ghrepo.info(function(err, body,status){
                if(err){
                    cout_fail++;
                }else{
                    rmap[gurl]=body;
                    cout_succ++;
                    console.log(body);
                }
                cb(null);
            });          
        }, function (err, result) {
            console.log(result);
    		callback(cout_succ,cout_fail, rmap);
        });
	},
	function(cout_succ,cout_fail, rmap, callback){
		// arg1 now equals 'three'
		callback(null, 'done');
	}
], function (cout_succ,cout_fail, rmap) {
   // result now equals 'done'
    console.log('succ:'+cout_succ+' fail:'+cout_fail);
    //write result to excel
var _headers = [ 'name', 'forks', 'open_issues', 'watchers','stargazers_count','size','created_at','updated_at','language','html_url']
var _data =[];
for(var x in rmap){
  //console.log("Key: %s, Value: %s", key, value);
  var item={};
  var val = rmap[x];
  for(var i=0; i<_headers.length; i++){
      var key = _headers[i];
      item[key]=val[key];
  }
  _data.push(item);
}

var headers = _headers
                .map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 }))
                .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
var data = _data
               .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) })))
              .reduce((prev, next) => prev.concat(next))
              .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
// 合并 headers 和 data
var output = Object.assign({}, headers, data);
// 获取所有单元格的位置
var outputPos = Object.keys(output);
// 计算出范围
var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
// 构建 workbook 对象
var wb = {
    SheetNames: ['mySheet'],
    Sheets: {
        'mySheet': Object.assign({}, output, { '!ref': ref })
    }
};
// 导出 Excel
XLSX.writeFile(wb, 'output.xlsx');    

});