let originalData = [];

/*
==================================================
EVENTOS
==================================================
*/

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
LEITURA ARQUIVO
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

  /*
  ==============================================
  EXCEL
  ==============================================
  */

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

  /*
  ==============================================
  CSV
  ==============================================
  */

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
PROCESSA CSV
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

/*
==================================================
PARSE CSV
==================================================
*/

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
ÁRVORE HIERÁRQUICA POR USUÁRIO
==================================================
*/

function renderTree(
  data,
  filter = ''
){

  const tree =
    document.getElementById('tree');

  tree.innerHTML = '';

  /*
  ==========================================
  AGRUPA POR USUÁRIO
  ==========================================
  */

  const users = {};

  data.forEach(item => {

    const searchable =
      JSON.stringify(item)
      .toLowerCase();

    if(
      filter &&
      !searchable.includes(filter)
    ) return;

    const login =
      item.UsuarioLogin || 'N/A';

    const nome =

      item.Nome ||
      item.NomeCompleto ||
      item.DisplayName ||
      '';

    if(!users[login]){

      users[login] = {
        nome:nome,
        tree:{}
      };

    }

    /*
    ======================================
    REMOVE SERVIDOR
    ======================================
    */

    let pasta =
      item.Pasta || '';

    pasta = pasta.replace(
      /^\\\\[^\\]+\\[^\\]+\\/i,
      ''
    );

    const folders =
      pasta
      .split(/\\\\|\\/g)
      .filter(Boolean);

    let current =
      users[login].tree;

    folders.forEach(folder => {

      if(!current[folder]){

        current[folder] = {
          __children:{},
          __items:[]
        };

      }

      current =
        current[folder]
        .__children;

    });

    /*
    ======================================
    ADICIONA ITEM NA ÚLTIMA PASTA
    ======================================
    */

    let ref =
      users[login].tree;

    folders.forEach((folder,index) => {

      if(index === folders.length - 1){

        ref[folder]
        .__items
        .push(item);

      }

      else{

        ref =
          ref[folder]
          .__children;

      }

    });

  });

  /*
  ==========================================
  CRIA NÓ PASTA
  ==========================================
  */

  function createFolderNode(
    folderName,
    folderData,
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
      `📁 ${folderName}`;

    wrapper.appendChild(title);

    const content =
      document.createElement('div');

    content.style.display =
      'none';

    /*
    ======================================
    SUBPASTAS
    ======================================
    */

    Object.entries(
      folderData.__children
    ).forEach(([childName,childData]) => {

      content.appendChild(

        createFolderNode(
          childName,
          childData,
          level + 1
        )

      );

    });

    /*
    ======================================
    NÃO MOSTRA PERMISSÕES NA RAIZ
    ======================================
    */

    const isRoot =
      level === 0;

    if(!isRoot){

      /*
      ====================================
      PERMISSÕES
      ====================================
      */

      folderData.__items
      .forEach(item => {

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

          <div class="grid">

            <div class="info">

              <span>Permissão</span>

              ${item.Permissao || ''}

            </div>

            <div class="info">

              <span>Grupo</span>

              ${item.GrupoPermissao || ''}

            </div>

          </div>

        `;

        content.appendChild(div);

      });

    }

    /*
    ======================================
    EXPANDIR
    ======================================
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

  /*
  ==========================================
  RENDERIZA USUÁRIOS
  ==========================================
  */

  Object.entries(users)
  .forEach(([login,userData]) => {

    const userWrapper =
      document.createElement('div');

    userWrapper.className =
      'folder';

    const userTitle =
      document.createElement('div');

    userTitle.className =
      'folder-title';

    userTitle.innerHTML = `

      👤 ${login}

      <div style="
        color:#94a3b8;
        font-size:12px;
        margin-top:4px;
        font-weight:normal;
      ">

        ${userData.nome}

      </div>

    `;

    userWrapper.appendChild(userTitle);

    const userContent =
      document.createElement('div');

    userContent.style.display =
      'none';

    Object.entries(userData.tree)
    .forEach(([folderName,folderData]) => {

      userContent.appendChild(

        createFolderNode(
          folderName,
          folderData
        )

      );

    });

    userTitle.addEventListener(
      'click',
      () => {

        userContent.style.display =

          userContent.style.display ===
          'none'

            ? 'block'

            : 'none';

      }
    );

    userWrapper.appendChild(userContent);

    tree.appendChild(userWrapper);

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