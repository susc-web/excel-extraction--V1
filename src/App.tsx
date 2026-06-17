import { useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import {
  FileSpreadsheet,
  Upload,
  X,
  ListChecks
} from "lucide-react";

type Row = Record<string, unknown>;

const OUTPUT_COLUMNS = [
  "Job ID",
  "Customer Name",
  "Customer Phone",
  "Address",
  "Installation Date",
  "CoES",
] as const;

type OutputCol = (typeof OUTPUT_COLUMNS)[number];


const mapping = {
  jobId: "Job Id",
  firstName: "Customer First Name",
  lastName: "Customer Surname",
  phone: "Customer Phone Number",
  address: "Customer address",
  suburb: "Customer Suburb",
  postcode: "Customer Postcode",
  state: "Customer State",
  installDate: "Date  Scheduled Date",
  product: "Primary Product 1 Brand",
  productModel: "Primary Product 1 Model",
  serial: "Serial",
  questionCol: "Question Block Type",
};



function formatCell(value: unknown) {

  if (
    value === undefined ||
    value === null ||
    value === ""
  ) return "";


  if(value instanceof Date){

    const day =
      String(value.getDate())
      .padStart(2,"0");

    const month =
      String(value.getMonth()+1)
      .padStart(2,"0");

    const year =
      String(value.getFullYear())
      .slice(-2);


    return `${day}/${month}/${year}`;
  }


  return String(value);
}

function formatPhone(value: unknown){

  if(!value) return "";


  let v =
    String(value)
    .replace(/\D/g,"");


  if(v.startsWith("61")){
    v =
    "0"+v.substring(2);
  }


  return v;
}



function joinNonEmpty(
 values:string[],
 sep:string
){

 return values
 .map(x=>x.trim())
 .filter(Boolean)
 .join(sep);

}



function App(){


const fileInputRef =
useRef<HTMLInputElement|null>(null);



const [fileName,setFileName] =
useState("");

const [rows,setRows] =
useState<Row[]>([]);


const [status,setStatus] =
useState("");



const reset = ()=>{

setFileName("");
setRows([]);
setStatus("");

if(fileInputRef.current){
 fileInputRef.current.value="";
}

};




const handleFile = async(file:File)=>{

try{

const buffer =
await file.arrayBuffer();


const wb =
XLSX.read(
buffer,
{
 type:"array",
 cellDates:true
}
);


const ws =
wb.Sheets[wb.SheetNames[0]];


const json =
XLSX.utils.sheet_to_json<Row>(
ws,
{
 defval:"",
 raw:false
}
);


console.log("Excel first row:", json[0]);


setRows(json);

setFileName(file.name);


setStatus(
`Loaded ${json.length} rows`
);


}
catch(error){

console.error(error);

setStatus("Failed to read Excel");

}

};



const onFileChange =
(
e:React.ChangeEvent<HTMLInputElement>
)=>{

const file =
e.target.files?.[0];


if(file)
handleFile(file);

};






const buildRowForJob = (
jobId:string,
jobRows:Row[]
):Record<OutputCol,string>=>{


const first =
jobRows[0] ?? {};



const findValue = (column:string)=>{

const row =
jobRows.find(
r =>
Object.keys(r)
.some(
k =>
k.trim().toLowerCase()
===
column.trim().toLowerCase()
)
);


if(!row)
return "";


const key =
Object.keys(row)
.find(
k =>
k.trim().toLowerCase()
===
column.trim().toLowerCase()
);


return key
?
row[key]
:
"";

};



const customerFirst =
findValue(mapping.firstName);


const customerLast =
findValue(mapping.lastName);

const customerName =
joinNonEmpty(
[
  formatCell(customerFirst),
  formatCell(customerLast)
],
" "
);



const get =
(key:keyof typeof mapping)=>{

const col =
mapping[key];

return formatCell(
first[col]
);

};








const phone =
formatPhone(
first[mapping.phone]
);




const address =
joinNonEmpty(
[
get("address"),
get("suburb"),
get("postcode"),
get("state")
],
", "
);




const groups =
new Map<string,{
product:string;
model:string;
serials:string[];
}>();




for(const row of jobRows){


const product =
formatCell(
row[mapping.product]
);


const model =
formatCell(
row[mapping.productModel]
);


const serial =
formatCell(
row[mapping.serial]
);



if(
!product &&
!model &&
!serial
) continue;



const key =
`${product}-${model}`;



if(!groups.has(key)){

groups.set(
key,
{
product,
model,
serials:[]
}
);

}



if(serial){

const g =
groups.get(key)!;


if(!g.serials.includes(serial)){
g.serials.push(serial);
}

}


}




const coes:string[]=[];



groups.forEach(g=>{


const title =
joinNonEmpty(
[
g.product,
g.model
],
" - "
);



if(title)
coes.push(title);



if(g.serials.length > 0){

  coes.push(`Outdoor : ${g.serials[0]}`);

  for(let i = 1; i < g.serials.length; i++){

    coes.push(`Indoor : ${g.serials[i]}`);

  }

}


});





return {

"Job ID":jobId,

"Customer Name":
customerName,


"Customer Phone":
phone,


"Address":
address,


"Installation Date":
get("installDate"),


"CoES":
coes.join("\n")

};

};
const createOutput = ()=>{


const jobMap =
new Map<string,Row[]>();



for(const row of rows){


const id =
String(
row[mapping.jobId] ?? ""
);



if(!id) continue;



if(!jobMap.has(id)){
jobMap.set(id,[]);
}



jobMap
.get(id)!
.push(row);

}




return Array.from(jobMap)
.map(
([id,data])=>
buildRowForJob(
id,
data
)
);


};






const downloadExcel = ()=>{


const output =
createOutput();



const ws =
XLSX.utils.json_to_sheet(
output,
{
header:[
...OUTPUT_COLUMNS
]
}
);




ws["!cols"] =
OUTPUT_COLUMNS.map(
c=>({

wch:
c==="CoES"
?80
:c==="Address"
?45
:25

})
);




const wb =
XLSX.utils.book_new();


XLSX.utils.book_append_sheet(
wb,
ws,
"Jobs"
);



XLSX.writeFile(
wb,
"formatted_output.xlsx"
);



setStatus(
`Exported ${output.length} jobs`
);


};






const preview =
useMemo(
()=>createOutput(),
[
rows
]
);




return (

<div className="
min-h-screen
bg-gray-100
p-6
">


<div className="
mx-auto
max-w-6xl
">


<div className="
mb-8
">


<div className="
flex
items-center
gap-2
text-blue-600
">


<FileSpreadsheet size={25}/>


<h1 className="
text-3xl
font-bold
text-black
">

Excel Job Formatter

</h1>


</div>


<p className="
mt-2
text-gray-600
">

Upload Excel → Preview → Download

</p>



{
status &&
<div className="
mt-3
rounded
bg-green-100
p-2
text-green-700
">

{status}

</div>
}


</div>





{
!fileName ? (



<div className="
rounded-xl
border-2
border-dashed
bg-white
p-14
text-center
">


<Upload
size={55}
className="
mx-auto
text-blue-600
"
/>



<h2 className="
mt-4
text-xl
font-semibold
">

Upload Excel File

</h2>




<input

ref={fileInputRef}

type="file"

accept="
.xlsx,.xls,.csv
"

className="
mt-5
"

onChange={onFileChange}

/>


</div>



):(




<div className="space-y-6">





<div className="
rounded-xl
bg-white
p-5
shadow
">


<div className="
flex
justify-between
">


<div>


<h2 className="
font-bold
">

{fileName}

</h2>



<p className="
text-sm
text-gray-500
">

{rows.length} rows


</p>


</div>



<button

onClick={reset}

className="
rounded
bg-red-100
px-3
py-2
text-red-600
"


>

<X
size={16}
className="inline"
/>

Remove

</button>



</div>




<div className="
mt-5
">


<button

onClick={downloadExcel}

className="
rounded
bg-blue-600
px-5
py-3
text-white
"


>


<ListChecks
size={18}
className="inline mr-2"
/>


Download Excel


</button>


</div>



</div>






{
preview.length > 0 && (



<div className="
overflow-auto
rounded-xl
bg-white
shadow
">


<table className="w-full border">


<thead className="bg-gray-200">


<tr>


{
OUTPUT_COLUMNS.map(c=>(

<th
key={c}
className="
border
p-3
text-left
"
>

{c}

</th>

))
}



</tr>


</thead>





<tbody>


{
preview.map(row=>(


<tr key={row["Job ID"]}>


{
OUTPUT_COLUMNS.map(c=>(


<td

key={c}

className="
border
p-3
align-top
whitespace-pre-line
"


>

{row[c]}


</td>


))

}



</tr>


))
}



</tbody>



</table>



</div>



)
}






</div>


)

}



</div>


</div>

);


}


export default App;
