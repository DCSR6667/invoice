const bodyParser = require('body-parser');
const express=require('express');
const app=express();
const fs=require('fs');
const path=require('path');
const puppeteer=require('puppeteer');
const XLSX=require('xlsx');
const numToWords=require('./numtoword');

const multer=require('multer');


app.use(bodyParser.urlencoded({extended:false}));
app.set('view engine','ejs');
app.set('views','views');
app.use(express.static(path.join(__dirname, "static")));
app.use(express.static(path.join(__dirname, "invoices")));
app.use(express.static(path.join(__dirname, "payslips")));

const fileStorage=multer.diskStorage({
    destination:(req,file,cb)=>
    {
        
        cb(null,'./xlfiles/');

    }
    ,
    filename:(req,file,cb)=>
    {
        cb(null,file.originalname);
    }
});




//app.use(multer({storage:fileStorage}).single('upload'));
app.use(multer({storage:fileStorage}).fields([{ name: 'upload', maxCount: 1 }, { name: 'invoice', maxCount: 1}]));




app.get('/',(req,res,next)=>
{
    res.render('index');

});

app.get('/invoice/:data',(req,res,next)=>
{
    var data=JSON.parse(req.params.data);
    res.render('invoice',{...data});

});

app.get('/payslip/:data',(req,res,next)=>
{
    var data=JSON.parse(req.params.data);
    res.render('payslip',{...data});

});

app.get('/sheet_invoices',(req,res,next)=>
{
    res.render('sheet_invoices.ejs');

});




app.post('/sheet_invoices_submit',(req,res,next)=>
{
    
    // var path1="/invoice/"+req.files.invoice[0].path;
    var path1=path.join(__dirname, "xlfiles", req.files.invoice[0].filename);

    


   

   const workbook=XLSX.readFile(path1);
    var worksheets={};


    for(const sheetname of workbook.SheetNames)
    {
        worksheets[sheetname]=XLSX.utils.sheet_to_json(workbook.Sheets[sheetname]);
    }

    var data=worksheets.Sheet1;

    
   

    for(var i in data)
    {
        for(var j in data[i])
        {
            if(data[i][j]=='default')
            {
                data[i][j]="";
            }
            
        }
        data[i]['words']=numToWords(data[i]['GRAND_TOTAL'])
    }



    


   


 res.redirect('/generate_invoices/'+JSON.stringify(data)); 


});



app.get('/generate_invoices/:data',async(req,res,next)=>
{
    

    var data=JSON.parse(req.params.data);
    var l=[];
    var foldername=data[0].MONTH+"_"+data[0].YEAR+'_invoices';

    var p=path.join(__dirname, "invoices", foldername);


        if(!fs.existsSync(p)) {
          console.log("not exists folder created");
          fs.mkdirSync(p);
        }  
     

   


   for(var i of data)
    {
        const filename=i.NAME+"_"+i.MONTH+i.DATE+"_invoice_ES_Search.pdf"; 
       
        
        try{

            const browser=await puppeteer.launch();
            const page=await browser.newPage();
            await page.goto(`${req.protocol}://${req.get('host')}`+"/invoice/"+JSON.stringify(i),
                {
                    waitUntil:"networkidle2"
                });
               

            await page.setViewport({width:2500,height:1300});

            const pdfn=await page.pdf({
                path:path.join(__dirname, "invoices", foldername,filename),          
                printBackground:true,
                format:"A4"
            });

           
            var url='/'+foldername+'/'+filename;
            l.push(url); 

            await browser.close();

            
        }
        catch(error)
        {
            console.log("something went wrong");

        } 


    } 

    

       res.render('view_list_invoices',{l:l}); 
    
});















app.get('/sheet_payslips',(req,res,next)=>
{
    res.render('sheet_payslips.ejs');

});

app.post('/sheet_payslips_submit',(req,res,next)=>
{
    
    
 
    var path1=path.join(__dirname, "xlfiles", req.files.upload[0].filename);
   
    const workbook=XLSX.readFile(path1);
    var worksheets={};


    for(const sheetname of workbook.SheetNames)
    {
        worksheets[sheetname]=XLSX.utils.sheet_to_json(workbook.Sheets[sheetname]);
    }

    var data=worksheets.Sheet1;

    /* var data=JSON.stringify(worksheets.Sheet1);
    data=JSON.parse(data); */

   

    for(var i in data)
    {
        for(var j in data[i])
        {
            if(data[i][j]=='default')
            {
                data[i][j]="";
            }
            
        }
        data[i]['words']=numToWords(data[i]['Credited_Salary'])
    }


    
    res.redirect('/generate_payslips/'+JSON.stringify(data));
  
    
    

});



app.get('/generate_payslips/:data',async(req,res,next)=>
{
    

    var data=JSON.parse(req.params.data);
    var l=[];
    var foldername=data[0].Month+"_"+data[0].Year+'_payslips';

    var p=path.join(__dirname, "payslips", foldername);

    

        if(!fs.existsSync(p)) {
          console.log("not exists folder created");
          fs.mkdirSync(p);
        }  
    

   



   for(var i of data)
    {
        const filename=i.Name+"_payslips_"+i.Month+"_"+i.Year+".pdf"; 
      

        
        try{

            const browser=await puppeteer.launch();
            const page=await browser.newPage();
            await page.goto(`${req.protocol}://${req.get('host')}`+"/payslip/"+JSON.stringify(i),
                {
                    waitUntil:"networkidle2"
                });

            await page.setViewport({width:2500,height:1300});

            const pdfn=await page.pdf({
                path:path.join(__dirname, "payslips", foldername,filename),
                printBackground:true,
                format:"A4"
            });
            var url='/'+foldername+'/'+filename;
            l.push(url);

            await browser.close();

            /*const pdfurl='/invoice/docs/'+filename;
            res.download(pdfurl,function(err)
            {
                if(err)
                {
                    console.log("something went wrong");
                }

            }); */

            
        }
        catch(error)
        {
            console.log("something went wrong");

        } 


    } 


 res.render('view_list_payslips',{l:l});
    
});

app.listen(3001,()=>
{
    console.log("server is running on port 3001");
});