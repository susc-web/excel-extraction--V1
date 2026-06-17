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


type InputKey =
  | "jobId"
  | "firstName"
  | "lastName"
  | "phone"
  | "address"
  | "suburb"
  | "postcode"
  | "state"
  | "installDate"
  | "product"
  | "productModel"
  | "serial"
  | "questionCol";



const INPUT_FIELDS: {
  key: InputKey;
  label: string;
  candidates: string[];
}[] = [

{
key:"jobId",
label:"Job ID",
candidates:[
"Job Id",
"Job ID",
"JobId"
]
},

{
key:"firstName",
label:"Customer First Name",
candidates:[
"Customer First Name"
]
},

{
key:"lastName",
label:"Customer Surname",
candidates:[
"Customer Surname",
"Customer  Surname"
]
},

{
key:"phone",
label:"Customer Phone Number",
candidates:[
"Customer Phone Number"
]
},

{
key:"address",
label:"Customer Address",
candidates:[
"Customer address",
"Customer Address"
]
},

{
key:"suburb",
label:"Customer Suburb",
candidates:[
"Customer Suburb"
]
},

{
key:"postcode",
label:"Customer Postcode",
candidates:[
"Customer Postcode"
]
},

{
key:"state",
label:"Customer State",
candidates:[
"Customer State"
]
},

{
key:"installDate",
label:"Scheduled Date",
candidates:[
"Scheduled Date",
"Date Scheduled"
]
},

{
key:"product",
label:"Primary Product Brand",
candidates:[
"Primary Product 1 Brand"
]
},

{
key:"productModel",
label:"Primary Product Model",
candidates:[
"Primary Product 1 Model"
]
},

{
key:"serial",
label:"Serial",
candidates:[
"Serial"
]
},

{
key:"questionCol",
label:"Question Type",
candidates:[
"Question Block Type",
"Question Block",
"Question"
]
}

];



const norm = (v:string)=>
v
.toLowerCase()
.replace(/[^a-z0-9]/g,"");



function findColumn(
headers:string[],
candidates:string[]
){

const target =
candidates.map(norm);


for(const h of headers){

if(
target.includes(norm(h))
){
return h;
}

}


return "";

}



function autoDetect(
headers:string[]
){

const result =
{} as Record<InputKey,string>;


for(const field of INPUT_FIELDS){

result[field.key] =
findColumn(
headers,
field.candidates
);

}


return result;

}



function formatCell(
value:unknown
){

if(
value === undefined ||
value === null ||
value === ""
){
return "";
}


if(value instanceof Date){

return value
.toISOString()
.split("T")[0];

}


return String(value);

}



function formatPhone(
value:unknown
){

if(!value)
return "";


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
separator:string
){

return values
.map(x=>x.trim())
.filter(Boolean)
.join(separator);

}
function App() {

const fileInputRef = useRef<HTMLInputElement | null>(null);

const [fileName,setFileName] =
useState<string>("");

const [rows,setRows] =
useState<Row[]>([]);

const [headers,setHeaders] =
useState<string[]>([]);

const [mapping,setMapping] =
useState<Record<InputKey,string>>(
autoDetect([])
);

const [status,setStatus] =
useState("");



const setCol = (
key:InputKey,
value:string
)=>
{
setMapping(prev=>({
...prev,
[key]:value
}));
};




const reset = ()=>{

setFileName("");
setRows([]);
setHeaders([]);
setMapping(autoDetect([]));
setStatus("");

if(fileInputRef.current){
fileInputRef.current.value="";
}

};




const handleFile = async(file:File)=>{


try{


const buffer =
await file.arrayBuffer();



const workbook =
XLSX.read(
buffer,
{
type:"array",
cellDates:true
}
);



const sheet =
workbook.Sheets[
workbook.SheetNames[0]
];


const json =
XLSX.utils
.sheet_to_json<Row>(
sheet,
{
defval:""
}
);



if(json.length===0){

setStatus(
"No rows found"
);

return;

}



const cols =
Object.keys(json[0]);



setRows(json);

setHeaders(cols);

setFileName(file.name);

setMapping(
autoDetect(cols)
);



setStatus(
`Loaded ${json.length} rows`
);



}
catch(err){

setStatus(
"Failed to read Excel"
);

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



const get = (
key:InputKey
)=>{

const col =
mapping[key];

return col
?
formatCell(first[col])
:
"";

};



const customerName =
joinNonEmpty(
[
get("firstName"),
get("lastName")
],
" "
);



const phone =
mapping.phone
?
formatPhone(
first[mapping.phone]
)
:
"";



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
new Map<
string,
{
product:string;
model:string;
serials:string[];
}
>();




for(const row of jobRows){


const product =
mapping.product
?
formatCell(row[mapping.product])
:
"";


const model =
mapping.productModel
?
formatCell(row[mapping.productModel])
:
"";


const serial =
mapping.serial
?
formatCell(row[mapping.serial])
:
"";



if(!product && !model && !serial)
continue;



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

const group =
groups.get(key)!;


if(!group.serials.includes(serial))
group.serials.push(serial);

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



g.serials.forEach(
s=>
coes.push(
`Serial Number: ${s}`
)
);


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







const downloadAll = ()=>{


const jobColumn =
mapping.jobId;



if(!jobColumn)
return;



const groups =
new Map<string,Row[]>();



for(const row of rows){


const id =
String(row[jobColumn] ?? "");



if(!id)
continue;



if(!groups.has(id))
groups.set(id,[]);



groups
.get(id)!
.push(row);


}




const output =
Array.from(groups)
.map(
([id,data])=>
buildRowForJob(
id,
data
)
);



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
?
50
:
25

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






const jobCount =
useMemo(()=>{


if(!mapping.jobId)
return 0;


const set =
new Set<string>();


rows.forEach(r=>{

const id =
String(
r[mapping.jobId] ?? ""
);


if(id)
set.add(id);

});


return set.size;


},[
rows,
mapping.jobId
]);





const preview =
useMemo(()=>{


if(!mapping.jobId)
return [];



const map =
new Map<string,Row[]>();



rows.forEach(r=>{


const id =
String(
r[mapping.jobId] ?? ""
);



if(!id)
return;



if(!map.has(id))
map.set(id,[]);



map.get(id)!.push(r);


});



return Array.from(map)
.map(([id,data])=>
buildRowForJob(id,data)
);


},[
rows,
mapping
]);
  return (

<div className="min-h-screen bg-gray-100 p-6">

<div className="mx-auto max-w-6xl">


<div className="mb-8">

<div className="flex items-center gap-2 text-blue-600">

<FileSpreadsheet size={24}/>

<h1 className="text-3xl font-bold text-black">
Excel Job Formatter
</h1>

</div>


<p className="mt-2 text-gray-600">
Upload your Excel file and create one formatted row per Job ID.
</p>


{status && (

<div className="mt-3 rounded bg-green-100 p-2 text-green-700">

{status}

</div>

)}

</div>





{!fileName ? (


<div

className="
rounded-xl
border-2
border-dashed
bg-white
p-12
text-center
"

>


<Upload
className="mx-auto text-blue-600"
size={50}
/>


<h2 className="mt-4 text-xl font-semibold">

Drop Excel file here

</h2>



<input

ref={fileInputRef}

type="file"

accept=".xlsx,.xls,.csv"

className="mt-5"

onChange={onFileChange}

/>


</div>


) : (



<div className="space-y-6">





<div className="rounded-xl bg-white p-5 shadow">


<div className="flex justify-between">


<div>


<h2 className="font-bold">

{fileName}

</h2>


<p className="text-sm text-gray-500">

{rows.length} rows · {headers.length} columns

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


<X size={16} className="inline"/>

Remove


</button>



</div>






<div className="mt-6 grid gap-4 md:grid-cols-2">


{INPUT_FIELDS.map(field=>(


<div key={field.key}>


<label className="text-sm font-medium">

{field.label}

</label>



<select


className="
mt-1
w-full
rounded
border
p-2
"


value={mapping[field.key]}


onChange={
e=>
setCol(
field.key,
e.target.value
)
}


>



<option value="">

Select column

</option>



{headers.map(h=>(


<option

key={h}

value={h}

>

{h}

</option>


))}


</select>



</div>


))}


</div>


</div>







<div className="rounded-xl bg-white p-5 shadow">


<div className="flex items-center justify-between">


<div>


<h2 className="font-bold">

Export

</h2>


<p className="text-sm text-gray-500">

{jobCount} jobs found

</p>


</div>




<button

onClick={downloadAll}

disabled={!mapping.jobId}

className="
rounded
bg-blue-600
px-5
py-3
text-white
disabled:opacity-40
"


>


<ListChecks
className="inline mr-2"
size={18}
/>


Download Excel


</button>



</div>


</div>







{preview.length > 0 && (


<div className="overflow-auto rounded-xl bg-white shadow">


<table className="w-full border">


<thead className="bg-gray-200">


<tr>


{OUTPUT_COLUMNS.map(c=>(


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


))}


</tr>


</thead>





<tbody>


{preview.map(row=>(


<tr key={row["Job ID"]}>


{OUTPUT_COLUMNS.map(c=>(


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


))}



</tr>


))}


</tbody>


</table>


</div>


)}




</div>


)}



</div>


</div>

);

}


export default App;
