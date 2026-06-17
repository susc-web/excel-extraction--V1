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
