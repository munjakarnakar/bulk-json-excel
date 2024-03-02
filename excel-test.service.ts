import * as Excel from 'exceljs';
import * as fs from 'fs';


const dummyJson = {
    lab_number: '637276363',
    reference_number: '637276363-104088-1',
    tube_colour_code: '104088',
    tube_color: 'GREY TOP',
    source: '1xView',
    precious_sample: false,
    registration_type: 1,
    registration_timestamp: '2023-03-02T12:19:17.000Z',
    registration_time: '2023-03-02T12:19:17.000Z',
    registration_source: '1xView',
    client_code: 'MD5496',
    current_registration_status: 'REGISTERED',
    invoice_code: 'C0000291',
    center_type: 'Pickup Point',
    center_name: 'SAIdPdOP',
    center_state: 'Andhra Pradesh',
    center_city: 'Guntur',
    network_type: 'FE TOWN',
    route_type: 'ROUTE',
    status: 'Registered',
    lab_number2: '637276363',
    reference_number2: '637276363-104088-1',
    tube_colour_code2: '104088',
    tube_color2: 'GREY TOP',
    source2: '1xView',
    precious_sample2: false,
    registration_type2: 1,
    registration_timestamp2: '2023-03-02T12:19:17.000Z',
    registration_time2: '2023-03-02T12:19:17.000Z',
    registration_source2: '1xView',
    client_code2: 'MD5496',
    current_registration_status2: 'REGISTERED',
    invoice_code2: 'C0000291',
    center_type2: 'Pickup Point',
    center_name2: 'SAIdPdOP',
    center_state2: 'Andhra Pradesh',
    center_city2: 'Guntur',
    network_type2: 'FE TOWN',
    route_type2: 'ROUTE',
    status2: 'Registered'
};

async function generateExcel(headers: any, data: any) {
    const file_name = `${data.length}-streamed-workbook.xlsx`;
    const reports_path = `${__dirname}/reports`;
    if (!fs.existsSync(reports_path)) {
        // If it doesn't exist, create it
        fs.mkdirSync(reports_path, { recursive: true });
        console.log({ descrisendingption: 'Excel-Test:Folder created successfully', jsonObject: { reports_path } });
    } else {
        console.log({ description: 'Excel-Test:Folder already exists', jsonObject: { reports_path } });
    }

    const file_path = `${reports_path}/${file_name}`;

    const options = {
        filename: file_path,
        useStyles: true,
        useSharedStrings: true
    };

    const workbook = new Excel.stream.xlsx.WorkbookWriter(options);

    console.log({ description: 'Excel-Test:Workbook created', jsonObject: {} });

    const worksheet = workbook.addWorksheet('data');

    worksheet.columns = headers;

    console.log({ description: 'Excel-Test: Added columns to work book', jsonObject: { headers } });

    console.log(`Writing data to workbook: **Columns ${headers.length}**`, `**Rows ${data.length}**`);

    for (let index = 0; index < data.length; index++) {
        console.log(index, data[index])
        worksheet.addRow(data[index]).commit();
    }

    await workbook.commit();

    console.log({ description: 'Excel-Test:Excel generated in temp path:', jsonObject: file_path });

    return { file_path, file_name };
};


generateExcel([{ header: 'Name', key: 'name' }, { header: 'Email', key: 'email' }, { header: 'Phone', key: 'phone' }], [{ name: 'Karnakar', phone: '9000423012', email: 'munja.karnakar@gmailc.com' }])