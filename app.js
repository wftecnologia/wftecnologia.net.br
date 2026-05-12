let originalData = [];

let permChart;
let riskChart;

Chart.defaults.color = '#ffffff';

document
.getElementById('search')
.addEventListener('input', function(){

  renderTree(
    originalData,
    this.value.toLowerCase()
  );

});

document
.getElementById('csvFile')
.addEventListener(
  'change',
  handleFile
);

document
.getElementById('generatePDF')
.addEventListener(
  'click',
  generatePDFReport
);

/*
==================================================
LEITURA
==================================================
*/

function handleFile(e){

  const file =
    e.target.files[0];

  if(!file) return;

  const reader =
    new FileReader();

  const fileName =
    file.name.toLowerCase();

  if(
    fileName.endsWith('.xlsx') ||
    fileName.endsWith('.xls')
  ){

    reader.onload = function(event){

      const data =
        new Uint8Array(
          event.target.result
        );

      const workbook =
        XLSX.read(
          data,
          { type:'array' }
        );

      const worksheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ];

      const jsonData =
        XLSX.utils.sheet_to_json(
          worksheet,
          { defval:'' }
        );

      loadData(jsonData);

    };

    reader.readAsArrayBuffer(file);

  }

  else{

    reader.onload = function(event){

      processCSV(
        event.target.result
      );

    };

    reader.readAsText(
      file,
      'UTF-8'
    );

  }

}

/*
==================================================
CSV
==================================================
*/

function processCSV(csv){

  const lines =
    csv
    .split('\n')
    .filter(
      l => l.trim() !== ''
    );

  const headers =
    parseCSVLine(lines[0]);

  const data = [];

  for(let i=1;i<lines.length;i++){

    const values =
      parseCSVLine(lines[i]);

    const obj = {};

    headers.forEach((h,index)=>{

      obj[h] =
        values[index] || '';

    });

    data.push(obj);

  }

  loadData(data);

}

function parseCSVLine(line){

  const result = [];

  let current = '';

  let insideQuotes = false;

  for(let char of line){

    if(char === '"'){

      insideQuotes =
        !insideQuotes;

    }

    else if(
      char === ',' &&
      !insideQuotes
    ){

      result.push(current);

      current = '';

    }

    else{

      current += char;

    }

  }

  result.push(current);

  return result.map(v =>

    v
    .replace(/^"|"$/g,'')
    .trim()

  );

}

/*
==================================================
LOAD
==================================================
*/

function loadData(data){

  originalData = data;

  renderStats(data);

  renderCharts(data);

  renderTree(data);

}

/*
==================================================
CARDS
==================================================
*/

function renderStats(data){

  const usuarios =
    new Set(
      data.map(
        x => x.UsuarioLogin
      )
    ).size;

  const grupos =
    new Set(
      data.map(
        x => x.GrupoPermissao
      )
    ).size;

  const pastas =
    new Set(
      data.map(
        x => x.Pasta
      )
    ).size;

  document
  .getElementById('stats')
  .innerHTML = `

    <div class="card">
      <span>Usuários</span>
      <h2>${usuarios}</h2>
    </div>

    <div class="card">
      <span>Grupos</span>
      <h2>${grupos}</h2>
    </div>

    <div class="card">
      <span>Pastas</span>
      <h2>${pastas}</h2>
    </div>

    <div class="card">
      <span>Registros</span>
      <h2>${data.length}</h2>
    </div>

  `;

}

/*
==================================================
GRÁFICOS
==================================================
*/

function renderCharts(data){

  if(permChart)
    permChart.destroy();

  if(riskChart)
    riskChart.destroy();

  const read =
    data.filter(x =>

      (x.Permissao || '')
      .toLowerCase()
      .includes('read')

    ).length;

  const modify =
    data.filter(x =>

      (x.Permissao || '')
      .toLowerCase()
      .includes('modify')

    ).length;

  const full =
    data.filter(x =>

      (x.Permissao || '')
      .toLowerCase()
      .includes('full')

    ).length;

  document
  .getElementById('riskRead')
  .innerText = read;

  document
  .getElementById('riskModify')
  .innerText = modify;

  document
  .getElementById('riskFull')
  .innerText = full;

  /*
  ==============================================
  PIZZA
  ==============================================
  */

  permChart = new Chart(

    document.getElementById(
      'permChart'
    ),

    {

      type:'doughnut',

      data:{

        labels:[
          'Read',
          'Modify',
          'Full'
        ],

        datasets:[{

          data:[
            read,
            modify,
            full
          ],

          backgroundColor:[
            '#22c55e',
            '#f59e0b',
            '#ef4444'
          ]

        }]

      },

      options:{

        responsive:true,

        maintainAspectRatio:false,

        plugins:{
          legend:{
            position:'bottom'
          }
        }

      }

    }

  );

  /*
  ==============================================
  RISCO
  ==============================================
  */

  riskChart = new Chart(

    document.getElementById(
      'riskChart'
    ),

    {

      type:'bar',

      data:{

        labels:[
          'Baixo',
          'Médio',
          'Crítico'
        ],

        datasets:[{

          label:'Permissões',

          data:[
            read,
            modify,
            full
          ],

          backgroundColor:[
            '#22c55e',
            '#f59e0b',
            '#ef4444'
          ],

          borderRadius:10

        }]

      },

      options:{

        responsive:true,

        maintainAspectRatio:false,

        indexAxis:'y',

        plugins:{
          legend:{
            display:false
          }
        }

      }

    }

  );

}

/*
==================================================
ÁRVORE
==================================================
*/

function renderTree(
  data,
  filter = ''
){

  const tree =
    document.getElementById(
      'tree'
    );

  tree.innerHTML = '';

  const treeData = {};

  data.forEach(item => {

    const searchable =
      JSON.stringify(item)
      .toLowerCase();

    if(
      filter &&
      !searchable.includes(filter)
    ) return;

    /*
    REMOVE SERVIDOR
    */

    let cleanedPath =
      item.Pasta || '';

    cleanedPath = cleanedPath.replace(
      /^\\\\[^\\]+\\[^\\]+\\/i,
      ''
    );

    const folders =
      cleanedPath
      .split(/\\\\|\\/g)
      .filter(Boolean);

    let current =
      treeData;

    folders.forEach(folder => {

      if(!current[folder]){

        current[folder] = {
          __items:[],
          __children:{}
        };

      }

      current =
        current[folder]
        .__children;

    });

    const lastFolder =
      folders[folders.length - 1];

    if(lastFolder){

      let folderRef =
        treeData;

      folders.forEach(folder => {

        folderRef =
          folderRef[folder];

        if(folder !== lastFolder){

          folderRef =
            folderRef.__children;

        }

      });

      folderRef
      .__items
      .push(item);

    }

  });

  /*
  ==============================================
  RENDERIZA
  ==============================================
  */

  function createFolderNode(
    name,
    data,
    level = 0
  ){

    const wrapper =
      document.createElement('div');

    wrapper.style.marginLeft =
      `${level * 18}px`;

    const title =
      document.createElement('div');

    title.className =
      'folder-title';

    title.innerHTML =
      `📁 ${name}`;

    wrapper.appendChild(title);

    const content =
      document.createElement('div');

    content.style.display =
      'none';

    /*
    ==========================================
    SUBPASTAS
    ==========================================
    */

    Object.entries(
      data.__children
    ).forEach(([childName, childData]) => {

      content.appendChild(

        createFolderNode(
          childName,
          childData,
          level + 1
        )

      );

    });

    /*
    ==========================================
    NÃO MOSTRA PERMISSÃO NA RAIZ
    ==========================================
    */

    const isRoot =
      level === 0;

    if(!isRoot){

      data.__items.forEach(item => {

        let permClass =
          'permission';

        const perm =
          (item.Permissao || '')
          .toLowerCase();

        if(
          perm.includes('full')
        ){

          permClass +=
            ' perm-full';

        }

        else if(
          perm.includes('modify')
        ){

          permClass +=
            ' perm-modify';

        }

        else if(
          perm.includes('read')
        ){

          permClass +=
            ' perm-read';

        }

        const div =
          document.createElement('div');

        div.className =
          permClass;

        div.innerHTML = `

          <strong>
            ${item.UsuarioLogin || 'N/A'}
          </strong>

          <div style="
            color:#94a3b8;
            margin-top:4px;
            font-size:12px;
          ">

            ${
              item.Nome ||
              item.NomeCompleto ||
              item.DisplayName ||
              ''
            }

          </div>

          <div class="grid">

            <div class="info">

              <span>Grupo</span>

              ${
                item.GrupoPermissao || ''
              }

            </div>

            <div class="info">

              <span>Permissão</span>

              ${
                item.Permissao || ''
              }

            </div>

          </div>

        `;

        content.appendChild(div);

      });

    }

    /*
    ==========================================
    EXPANDIR
    ==========================================
    */

    title.addEventListener(
      'click',
      () => {

        content.style.display =

          content.style.display ===
          'none'

            ? 'block'

            : 'none';

      }
    );

    wrapper.appendChild(content);

    return wrapper;

  }

  Object.entries(treeData)
  .forEach(([name, data]) => {

    tree.appendChild(

      createFolderNode(
        name,
        data
      )

    );

  });

}

/*
==================================================
PDF
==================================================
*/

function generatePDFReport(){

  const login = prompt(
    'Digite o login do colaborador:'
  );

  if(!login) return;

  const filtered =
    originalData.filter(

      x =>

      (x.UsuarioLogin || '')
      .toLowerCase()
      .includes(
        login.toLowerCase()
      )

    );

  if(filtered.length === 0){

    alert(
      'Usuário não encontrado.'
    );

    return;

  }

  const { jsPDF } =
    window.jspdf;

  const doc =
    new jsPDF('landscape');

  doc.setFontSize(18);

  doc.text(
    'Relatório de Permissões',
    14,
    20
  );

  doc.setFontSize(11);

  doc.text(
    `Colaborador pesquisado: ${login}`,
    14,
    30
  );

  const uniqueRows =
    new Map();

  filtered.forEach(item => {

    const nome =

      item.Nome ||
      item.NomeCompleto ||
      item.DisplayName ||
      item.UsuarioLogin ||
      '';

    let pasta =
      item.Pasta || '';

    pasta = pasta.replace(
      /^\\\\[^\\]+\\[^\\]+\\/i,
      ''
    );

    const permissao =
      item.Permissao || '';

    const key =
      `${nome}|${pasta}|${permissao}`;

    if(!uniqueRows.has(key)){

      uniqueRows.set(

        key,

        [
          nome,
          pasta,
          permissao
        ]

      );

    }

  });

  const rows =
    Array.from(
      uniqueRows.values()
    );

  doc.autoTable({

    startY:40,

    head:[[
      'Nome do colaborador',
      'Diretório',
      'Permissão'
    ]],

    body:rows,

    styles:{
      fontSize:9,
      cellPadding:4
    },

    headStyles:{
      fillColor:[37,99,235]
    },

    alternateRowStyles:{
      fillColor:[240,240,240]
    },

    columnStyles:{
      0:{cellWidth:70},
      1:{cellWidth:160},
      2:{cellWidth:40}
    }

  });

  doc.save(
    `Relatorio_${login}.pdf`
  );

}