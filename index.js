const ExcelJS = require('exceljs');
const { MongoClient, ServerApiVersion } = require("mongodb");

const uri = "mongodb://localhost:27017/sample";

const client = new MongoClient(uri, {
    serverApi: {
        version: ServerApiVersion.v1,
        strict: true,
        deprecationErrors: true,
    }
}
);

async function start() {
    try {

        await client.connect();
        let db = client.db("sample")
        const collection = db.collection("users");
        console.log("Pinged your deployment. You successfully connected to MongoDB!");

        

        

        let data = []
        const workbook = new ExcelJS.stream.xlsx.WorkbookReader('./sample1.xls');

        let count = 0
        for await (const worksheetReader of workbook) {
            for await (const row of worksheetReader) {
                count++;
                data.push(
                    {
                        insertOne: {
                            document: {
                                name:row.values['1']
                            },
                        },
                    },
                    
                )
                if (count % 10000) {
                    console.log(`${count} processed`);
                    const used = process.memoryUsage().heapUsed / 1024 / 1024;
                    // console.log('\033c');
                    console.log(`The script uses approximately ${Math.round(used * 100) / 100} MB`);

                    // await saveInDb()
                    const result = await collection.bulkWrite(data);
                    data=[]
                }
                // console.log(Object.values(row.values));
                // console.log(row.values['1']);


            }
        }

    } catch (error) {
        console.log(error);
    }
}



start()