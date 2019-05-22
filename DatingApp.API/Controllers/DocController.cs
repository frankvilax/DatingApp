using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using DatingApp.API.Data;
using DatingApp.API.Dtos;
using DatingApp.API.Models;
using HandlebarsDotNet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using RazorLight;

namespace DatingApp.API.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class DocController : ControllerBase
    {
        private readonly DataContext _context;

        public DocController(DataContext context)
        {
            _context = context;
        }

        public List<int> AllIndexesOf(string str, string value) {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("the string to find may not be empty", "value");
            List<int> indexes = new List<int>();
            for (int index = 0;; index += value.Length) {
                index = str.IndexOf(value, index);
                if (index == -1)
                    return indexes;
                indexes.Add(index);
            }
        }

        public string ReplaceRow(string html, int start, int end, string value){
            
            string result = html.Substring(start+4, end - start - 4);
            html = html.Replace(result, value);

            return html;
        }


        public string GenerateListTemplate(){
            
                    string source =
            @"<h2>Names</h2>
            {{#names}}
            {{> user}}
            {{/names}}";

            string partialSource =
            @"<strong>{{name}}</strong>";

            Handlebars.RegisterTemplate("user", partialSource);

            var template = Handlebars.Compile(source);

            var data = new {
                names = new [] 
                {
                    new {
                        name = "Karen"
                    },
                    new {
                        name = "Jon"
                    },

                    new {
                        name = "Nelson"
                    },
                    new {
                        name = "Amilcar"
                    }
                }
            };

            var result = template(data);

            return result;
        }

        public string GenerateTableTemplate(){
            string source = @"
            {{#if Bands}}
            <table>
            <tr>
                <th>Band Name</th>
                <th>Date</th>
                <th>First Album Name</th>
                <th>Second Album Name</th>
            </tr>
            {{#each Bands}}
                <tr>
                    <td>{{Name}}</td>
                    <td>{{Date}}</td>
                    <td>{{Albums.0.Name}}</td>
                    <td>{{Albums.1.Name}}</td>
                </tr>
            {{/each}}
        </table>
        {{else}}
            <h3>There are no concerts coming up.</h3>
        {{/if}}";
            
            var template = Handlebars.Compile(source);
            
            var data = new {
            bands = new [] {
                new {
                    name = "Karen",
                    date = "Aug 14th, 2013",
                    albums = new [] {
                new {
                    name = "Album name1"
                },
                new {
                    name = "Album name5"
                }}
                },
                new {
                    name = "Jon",
                    date = "Aug 15th, 2014",
                    albums = new [] {
                new {
                    name = "Album name1"
                },
                new {
                    name = "Album name6"
                }}
                },

                new {
                    name = "Nelson",
                    date = "Aug 16th, 2015",
                    albums = new [] {
                new {
                    name = "Album name2"
                },
                new {
                    name = "Album name7"
                }
                }
                },
                new {
                    name = "Amilcar",
                    date = "Aug 17th, 2016",
                    albums = new [] {
                new {
                    name = "Album name3"
                },
                new {
                    name = "Album name4"
                }
                }
                }
            }
            };


            var result = template(data);

            return result;

        }

        public string GenerateValueTemplate(){
            string source = @"
            {{#if Values}}
            <table>
            <tr>
                <th>Value Id</th>
                <th>Value Name</th>
            </tr>
            {{#each Values}}
                <tr>
                    <td>{{Id}}</td>
                    <td>{{Name}}</td>
                </tr>
            {{/each}}
        </table>
        {{else}}
            <h3>There are no values coming up.</h3>
        {{/if}}";
            
            var template = Handlebars.Compile(source);
            
            var data = new {
                values = new Value[] 
                {
                    new Value{
                        Id = 1,
                        Name = "Aug 14th, 2013"
                    },
                new Value{
                        Id = 1,
                        Name = "Aug 14th, 2013"
                    }
                }
            };
            


            var result = template(data);

            return result;

        }

        public List<Product> LoadJson()
        {
            using (StreamReader r = new StreamReader("../../../" +"ListPublicFullProduct.json"))
            {
                string json = r.ReadToEnd();
                List<Product> products = JsonConvert.DeserializeObject<List<Product>>(json);

                return products;
            }

        }

        [HttpGet]
        public async Task<IActionResult> CreateDocument(){
            

            var dc = await _context.Docs.FirstOrDefaultAsync(d => d.Id == 1);
           
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello World!");  

            builder.InsertHtml(GenerateTableTemplate());

            string dataDir  = "../../../Docs/" + "TestDB5.docx";
            // Save the document to disk.
            doc.Save(dataDir);


            string dataDir2 = "../../../Docs/" + "TestDB5.pdf";
            doc.Save(dataDir2);


            string html_template = dc.Template;

            // Generate Excel Table
            System.IO.File.WriteAllText(@"../../../Docs/TestDB5.xls", GenerateTableTemplate());

            // Calculate index of begin of the row
            List<int> row_begin = AllIndexesOf(html_template, "<td>");

            // Calculate index of end of the row
            List<int> row_end = AllIndexesOf(html_template, "</td>");
            
            ReplaceRow(html_template, 69, 77, "John" );

            
            return Ok(GenerateValueTemplate());
        }



        [HttpGet("{id}", Name = "GetTemplate")]

        public async Task<ActionResult> GetTemplate(int id){
            
            var template = await _context.Docs.FirstOrDefaultAsync(d => d.Id == id);

            return Ok(template);
        }

        [HttpPost("add")] 
        public async Task<ActionResult> AddTemplate(Doc doc){


            await _context.Docs.AddAsync(doc);
            await _context.SaveChangesAsync();

            return Ok(doc);
        }


        [HttpPost("create/{type}")]
        
        public async Task<ActionResult> CreateDoc(UserForRegisterDto[] users, string type){
            

            var temp = await _context.Docs.FirstOrDefaultAsync(d => d.Id == 2);
            string source = temp.Template;
            

            string test = @"<table> <caption>A complex table</caption> <thead> <tr> <th >Invoice #123456789</th> <th>14 January 2025 </tr> <tr> <td> <strong>Pay to:</strong><br> Acme Billing Co.<br> 123 Main St.<br> Cityville, NA 12345 </td> <td colspan=> <strong>Customer:</strong><br> John Smith<br> 321 Willow Way<br> Southeast Northwestershire, MA 54321 </td> </tr> </thead> <tbody> <tr> <th>Name / Description</th> <th>Qty.</th> <th>@</th> <th>Cost</th> </tr> <tr> <td>Paperclips</td> <td>1000</td> <td>0.01</td> <td>10.00</td> </tr> <tr> <td>Staples (box)</td> <td>100</td> <td>1.00</td> <td>100.00</td> </tr> </tbody> <tfoot> <tr> <th >Subtotal</th> <td> 110.00</td> </tr> <tr> <th >Tax</th> <td> 8% </td> <td>8.80</td> </tr> <tr> <th >Grand Total</th> <td>$ 118.80</td> </tr> </tfoot> </table> 
            <style> table, th, td {
  border: 1px solid black;
} </style>";

            var data = new { users };
            var template = Handlebars.Compile(source);
            var result = template(data);

            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // builder.Writeln("Hello World!\n");  

            builder.InsertHtml(result);
            
            if (type == "xls"){
                System.IO.File.WriteAllText(@"../../../Docs/Test.xls", result);
            }
            
            else if(type == "doc" || type == "pdf") {
                string dataDir  = "../../../Docs/" + "Test." +type;

                // Save the document to disk.
                doc.Save(dataDir);
            }
            else {
                return BadRequest("Invalid Format!");
            }
        

            return Ok();
        }


        [HttpGet("test1")]
        public void TestHelper(){
            HandlebarsBlockHelper _stringEqualityBlockHelper = (TextWriter output, HelperOptions options, dynamic context, object[] arguments) => {
            
                    if (arguments.Length != 2)
                    {
                        throw new HandlebarsException("{{StringEqualityBlockHelper}} helper must have exactly two argument");
                    }
                    
                    string left = arguments[0] as string;
                    
                    string right = arguments[1] as string;
                    
                    if (left == right)
                    {
                        options.Template(output, null);
                    }
                    
                    else
                    {
                        options.Inverse(output, null);
                    }

            };
            
            Handlebars.RegisterHelper("StringEqualityBlockHelper", _stringEqualityBlockHelper);
            
            Dictionary<string, string> animals = new Dictionary<string, string>() {
                {"Fluffy", "cat" },
                {"Fido", "dog" },
                {"Chewy", "hamster" }
            };
            
            string template = "{{#each @value}}The animal, {{@key}}, {{StringEqualityBlockHelper @value 'dog'}}is a dog{{else}}is not a dog{{/StringEqualityBlockHelper}}.\r\n{{/each}}";
            
            Func<object, string> compiledTemplate = Handlebars.Compile(template);
            
            string templateOutput = compiledTemplate(animals);

            Console.WriteLine(templateOutput);

            /* Would render
            The animal, Fluffy, is not a dog.
            The animal, Fido, is a dog.
            The animal, Chewy, is not a dog.
            */
        }

        [HttpGet("test2")]
        public void TestHelper2(){
            
            var basicTestTemplate = "{{mycustomhelper Person.Address.City}}";
            var blockTestTemplate = "{{#mycustomblock Person.Age }} Bye bye world {{else}}is not a dog{{/mycustomblock}}";

            Handlebars.RegisterHelper (
                "mycustomhelper", 
                (writer, context, arguments) => writer.Write (arguments [0]));

            Handlebars.RegisterHelper (
                "mycustomblock", (writer, options, context, arguments) => {
                    var arg = arguments[0];
        
            
                if ((int) arg == 30)
                {
                    options.Template(writer, null);
                }
                
                else
                {
                    options.Inverse(writer, null);
                }
                    });

                var data = new {
                    Person = new {
                        FirstName = "Bob",
                        Age = 30,
                        Address = new {
                            City = "New York"
                        }
                    }
            };

            var basicTest = Handlebars.Compile (basicTestTemplate);
            var blockTest = Handlebars.Compile (blockTestTemplate);

            Console.WriteLine (basicTest (data));
            Console.WriteLine (blockTest (data));
            Console.Read ();
    }


        [HttpPost("createRazor/{type}")]
        
        public async Task<ActionResult> CreateDocRazor(ItemToReportDto[] items, string type){
            
            var temp = await _context.Docs.FirstOrDefaultAsync(d => d.Id == 10);
            string template = temp.Template;

            var engine = new RazorLightEngineBuilder()
              .UseMemoryCachingProvider()
              .Build();

            var data = new { items };
            
            string result = await engine.CompileRenderAsync("templateKey", template, data);

            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // builder.Writeln("Hello World!\n");  

            builder.InsertHtml(result);

            // Create a new memory stream.
            MemoryStream outStream = new MemoryStream();

            // Save the document to stream.
            doc.Save(outStream, SaveFormat.Doc);

            // Convert the document to byte form.
            byte[] docBytes = outStream.ToArray();


            // From bytes to Doc/Pdf/Xls
            System.IO.File.WriteAllBytes(@"../../../Docs/Test." + type, docBytes);


            /* if (type == "xls"){
            System.IO.File.WriteAllText(@"../../../Docs/Test." + type, result);
            }*/
            
            /* else if(type == "doc" || type == "pdf") {
                string dataDir  = "../../../Docs/" + "Test." +type;

                // Save the document to disk.
                doc.Save(dataDir);
            }
            else {
                return BadRequest("Invalid Format!");
            }*/
            

            return Ok();
        }

}
}