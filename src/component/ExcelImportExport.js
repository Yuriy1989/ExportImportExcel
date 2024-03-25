import { useState } from 'react';
import {read, utils, writeFile} from 'xlsx';
import { Space, Table, Tag } from 'antd';
import { moment } from 'moment';

const ExcelImportExport = () => {
    const [fileName, setFileName] = useState()
    const [dataTable, setDataTable] = useState([]);

    const handleImport = async (e) => {
        const file = e.target.files[0];
        setFileName(file.name);
        const data = await file.arrayBuffer();
        const wb = read(data, {cellText:false, cellDates: true});

        const ws = wb.Sheets[wb.SheetNames[0]];
        const jsonData = utils.sheet_to_json(ws, {dateNF: 'dd"."mm"."yyyy', raw: false});

        // const sortingDate = jsonData.sort((a, b) =>
        //     moment(b['Дата регистрации'], 'DD.MM.YYYY').diff(moment(a['Дата регистрации'], 'DD.MM.YYYY'))
        // );

        // const sortingDate = (jsonData) => {
        //     jsonData.map((item, index) => {
        //         console.log(item['Дата регистрации']);
        //     })

        // }

        // sortingDate(jsonData);

        setDataTable(jsonData);
        // console.log(jsonData);
        // jsonData.sort((a, b) => new Date(a['Дата регистрации']) – new Date(b['Дата регистрации']));
        // jsonData.sort((a, b) => new Date(a['Дата регистрации']).getTime() - new Date(b['Дата регистрации']).getTime());
        console.log(jsonData);
        // console.log(sortingDate);
    }

    const sorting = (data) => {
        // const t = data.sort((a, b) => (new Date(a['Дата регистрации'])) – (new Date(b['Дата регистрации'])));
        // console.log(t);/
        // return t;
    } 

    function handleExport() {
        const data = dataTable[0]['Дата регистрации'];
        console.log('data', data);
        const headings = [[
            'id',
            'name',
            'date'
        ]];

        const wb = utils.book_new();
        const ws = utils.json_to_sheet([]);
        utils.sheet_add_aoa(ws, headings);
        utils.sheet_add_json(ws, dataTable, { origin: 'A1', skipHeader: false });
        utils.book_append_sheet(wb, ws, 'Report');
        writeFile(wb, 'Report.xlsx');
    }
    
    return (
        <>
            <div>
                <input type='file' name='file' required onChange={handleImport}/>
            </div>
            <div>
                {fileName && (
                    <p>Имя файла: <span> {fileName} </span></p>
                )}
                <p>Загружено {dataTable.length ? dataTable.length : 0 } строк</p>
                {/* <Table dataSource={dataSource} columns={columns} />; */}
                <table>
                    <thead>
                        <tr>
                            {
                                dataTable.length ?
                               Object.keys(dataTable[0]).map((item, index) => (
                                    <th key={index}>{item}</th>
                                ))
                                    :
                                    <th>No Data</th>
                            }
                        </tr>
                    </thead>
                    <tbody>
                        {
                            dataTable.length?
                                dataTable.map((item, index) => {
                                    return (
                                        <tr key={index}>
                                            <td>{item['Дата регистрации']}</td>
                                        </tr>
                                    )
                                })
                                :
                            <tr><td>No Data</td></tr>
                        }
                    </tbody>
                </table>
            </div>
            <div>
                <button type='button' onClick={handleExport}>Сформировать отчет за прошедшую неделю</button>
            </div>
            <div>
                <button type='button' onClick={sorting(dataTable)}>Отсортировать</button>
            </div>
        </>
    )
}
 export default ExcelImportExport;
