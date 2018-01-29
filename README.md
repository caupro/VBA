# VBA
<html>
    <head>
    <title>Busca - DCCE - Versão 1.0</title>
    <meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type">

    <!-- Adicionando JQuery -->
 <script>
inserir jquery
</script>

 <!-- Adicionando Javascript -->
<script type=text/javascript>
    $(document).ready(function() {
	            function limpa_formulário_cnpj() {
                // Limpa valores do formulário de cep.
                $("#capital_social").val("");
                $("#ultima_atualizacao").val("");
                $("#nome").val("");
                $("#fantasia").val("");
                $("#tipo").val("");
                $("#situacao").val("");
                $("#abertura").val("");
                $("#rua").val("");
                $("#complemento").val("");
                $("#numero").val("");
                $("#bairro").val("");
                $("#cep").val("");
                $("#municipio").val("");
                $("#uf").val("");
                $("#data_situacao").val("");
				$("#atividade").val("");

            }
    $("#cnpj").blur(function() {
                var cnpj = $(this).val().replace(/\D/g, '');
				if (cnpj != "") {
					var validacnpj = /^[0-9]{14}$/;
						if(validacnpj.test(cnpj)){
						
						$("#capital_social").val("***");
                        $("#ultima_atualizacao").val("***");
                        $("#nome").val("***");
                        $("#fantasia").val("***");
                        $("#tipo").val("***");
                        $("#situacao").val("***");
                        $("#abertura").val("***");
                        $("#rua").val("***");
                        $("#complemento").val("***");
                        $("#numero").val("***");
                        $("#bairro").val("***");
                        $("#cep").val("***");
                        $("#municipio").val("***");
                        $("#uf").val("***");
                        $("#data_situacao").val("***");
						$("#atividade").val("***");

                //Verifica se campo cep possui valor informado.
                //Valida o formato do CNPJ.
					
        /* sem cache */
        $.ajaxSetup({ cache: false });
        $.getJSON("http://www.receitaws.com.br/v1/cnpj/" + cnpj + "/?json/?callback=meu_callback", function(info){
		   //$.getJSON("https://www.receitaws.com.br/v1/cnpj/" + cnpj + "/?json/?callback=meu_callback", function(data){
		  if (!("erro" in info)){
		  console.log (info);
            var html = [];

								$("#capital_social").val(info.capital_social);
                                $("#ultima_atualizacao").val(info.ultima_atualizacao);
                                $("#nome").val(info.nome);
                                $("#fantasia").val(info.fantasia);
                                $("#tipo").val(info.tipo);
                                $("#situacao").val(info.situacao);
                                $("#abertura").val(info.abertura);
                                $("#rua").val(info.logradouro);
                                $("#complemento").val(info.complemento);
                                $("#numero").val(info.numero);
                                $("#bairro").val(info.bairro);
                                $("#cep").val(info.cep);
                                $("#municipio").val(info.municipio);
                                $("#uf").val(info.uf);
                                $("#data_situacao").val(info.data_situacao);
								$("#atividade").val(info.atividade_principal[0].text);
								}
							else {
									//CEP pesquisado não foi encontrado.
									limpa_formulário_cnpj();
									alert("CNPJ não encontrado.");
								}
								/* loop de array */
								//$.each(info.atividade_principal, function(index, d){
								//	html.push("Atividade: ", d.text, ", ",/* cada , le um campo adicional */
								//			  "Cnae: ", d.code, ", ",
								//			  "<br>");
								//});
								$.each(info.atividades_secundarias, function(index, d){
									html.push("Atividade Secundaria: ", d.text, " - ",
											  "Cnae: ", d.code, 
											  "<br>");
								});
							$.each(info.qsa, function(index, d){
									html.push("NOME: ", d.nome, " - ",
											  "FUNÇÃO: ", d.qual, " ",
											  "<br>");
								});
											
								$("#div01").html(html.join('')).css("background-color" );
		});
		}
					else {
					//CNPJ é inválido.
					limpa_formulário_cnpj();
					alert("Formato de CNPJ inválido.");
					}
					}	
					else {
                    //CNPJ sem valor, limpa formulário.
					limpa_formulário_cnpj();
					}
    });
});

    </script>
	
    <style type="text/css">
        input[type="text"]{
        padding:1px;
        border:thin solid #ccc;
        color:#5d5d5d;
        }
        input[type="date"]{
        width:auto;
        padding:6px;
        border:thin solid #ccc;
        color:#5d5d5d;
        }
        input[type="text"]:focus{
        color:#000;
        font-style:normal;
        }
        input[type="date"]:focus{
        color:#000;
        font-style:normal;
        font-weight:bold;
        }
        input[type="submit"]{
        cursor: pointer;
        font-weight: bold;
        padding:6px;
        border:thin solid #fff;
        background:#069;
        color:#fff;
        }
        input[type="submit"]:hover]{
        background:#09f;
        color:#fff;
        }
        #btn382{
        cursor: pointer;
        font-weight: bold;
        padding:5px;
        border:thin solid #fff;
        background:#069;
        color:#fff;
        }
        #run{
        cursor: pointer;
        font-weight: bold;
        padding:4px;
        border:thin solid #fff;
        background:#069;
        color:#fff;
        }
        #reset{
        cursor: pointer;
        font-weight: bold;
        padding:4px;
        border:thin solid #fff;
        background:#069;
        color:#fff;
        }
        form,
        div {
          display: auto;
          padding: 5px;
        }
        form {
        margin: auto;
        background-color: #939399;
          display: auto;
          padding: 5px;
          padding-top: 5px;
          padding-bottom:150%;
          padding-left: 5px;

        }
        div {
        color: #FFF;
        text-align: left;
        height: 1px;
        }
        input {
        display: inline-block;
        padding-left: 3px;
        }
        #nome {
        display: inline-block;
        padding-left: 3px;
        }
        #fantasia {
          display: inline-block;
          padding-left: 3px;
        }
        #tipo {
          display: inline-block;
          padding-left: 3px;
        }
        #abertura {
          display: inline-block;
          padding-left: 3px;
        }
        #atividade {
        display: inline-block;
        padding-left: 3px;
        }
        #rua {
          display: inline-block;
          padding-left: 3px;
        }
        #complemento {
          display: inline-block;
          padding-left: 3px;
        }
        #numero {
          display: inline-block;
          padding-left: 3px;
        }
        #bairro {
          display: inline-block;
          padding-left: 3px;
        }
        #cep {
          display: inline-block;
          padding-left: 3px;
        }
        #municipio {
          display: inline-block;
          padding-left: 3px;
        }
        #uf {
          display: inline-block;
          padding-left: 2px;
        }
        #qsa {
          display: inline-block;
          padding-left: 3px;
        }
        #qsa1 {
          display: inline-block;
          padding-left: 3px;
        }
        #atividade {
        display: inline-block;
        padding-left: 3px;
        }
        #ultima_atualizacao {
        display: inline-block;
        padding-left: 3px;
        }
        iframe {
          width:44%;
          height:44%;
          position: absolute;
          top: 15px;
          right: 8px;
          /*padding:8px;*/
          margin: 8px;
          margin-top: 8px;
          margin-right: 15px;
          margin-bottom: 7px;
          margin-left: 8px;
          border:thin solid #ccc;
          }
	</style>
</head>

 <body>
  <!-- Inicio do formulario -->
    <form method="get" action="">
					<font face="Tahoma" style="font-size: 10pt;">
					<div>CNPJ: <input name="cnpj" type="text" id="cnpj" value="" size="18" maxlength="20""/></div><br />
					<br />
					<div>Nome: <input name="nome" type="text" id="nome" size="70" /></div><br />
					<div>Nome Fantasia: <input name="fantasia" type="text" id="fantasia" size="70" /></div><br />
					<div>Tipo: <input name="tipo" type="text" id="tipo" size="6" /> Situação: <input name="situacao" type="text" id="situacao" size="6" /> Fundação: <input name="abertura" type="text" id="abertura" size="8	" /> Data Situação: <input name="data_situacao" type="text" id="data_situacao" size="8" /></div><br />
					<div>Rua: <input name="rua" type="text" id="rua" size="40" /> Número: <input name="numero" type="text" id="numero" size="4"/> <input type="button" id="run" value="Google" /></div><br />
					<div>Complemento: <input name="complemento" type="text" id="complemento" size="61" /></div><br />
					<div>Bairro: <input name="bairro" type="text" id="bairro" size="20" /></div><br />
					<div>Municipio: <input name="municipio" type="text" id="municipio" size="32" /> CEP: <input name="cep" type="text" id="cep" size="10" /> Estado: <input name="uf" type="text" id="uf" size="2" /></div><br />
					<div>Atividade Principal: <input name="ultima_atualizacao" type="text" id="atividade" size="75" /></div><br />
					<div>Capital Social: <input name="capital_social" type="text" id="capital_social" size="20"/> Data Base: <input name="ultima_atualizacao" type="text" id="ultima_atualizacao" size="24" /></div><br />

    </form>
	<button id="reset" type="submit" >Limpar </button> 
	<script type=text/javascript>
					var nome = document.querySelector("#nome"); //O querySelector pega o valor do campo, neste caso por ID(#) - e ele só aceita string, por isso as aspas
					var rua = document.querySelector ("#rua");
					var municipio = document.querySelector ("#municipio");
					var numero = document.querySelector ("#numero");
					var buscatx = document.querySelector("#run"); //Variavel para o botao
					var closetx = document.querySelector("#reset"); //Variavel para o botao

					buscatx.addEventListener("click", function openWin(){ 
					//Quando o botao 'adicionar' for clicado ele mostrara a mensagem com os valores do form
					gwindow=window.open("https://www.google.com.br/?gws_rd=cr,ssl&ei=-XYkWc6CDYGewgSkhbfAAw#q=%22"+ nome.value + "%22+ crime OR lavagem OR sonegação OR corrupção OR desvio OR cartel OR doleiro OR operação OR policia OR condena OR ilícito OR prostituição OR esquema OR trafico OR irregularidade OR sanção OR terrorismo OR estelionato OR fraude OR delito OR investigação OR escravo OR prisão OR improbidade OR ocultação OR evasão OR quadrilha");
					//window.open("https://br.search.yahoo.com/search;_ylc=X3oDMTFiN25laTRvBF9TAzIwMjM1MzgwNzUEaXRjAzEEc2VjA3NyY2hfcWEEc2xrA3NyY2h3ZWI-?p="+ nome.value);
					//window.open("https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21");
					// window.open("http://www.receita.fazenda.gov.br/PessoaJuridica/CNPJ/cnpjreva/cnpjreva_solicitacao2.asp");
					//window.open("http://www4.net.bradesco.com.br/rest/consultareceitafederal.consultascpfcnpj.consultascpfcnpj.jsf");
					swindow=window.open("https://www.google.com.br/maps/place/"+ rua.value + " " + numero.value + " " + municipio.value + "/");
					//winSimple=window.open("http://www4.net.bradesco.com.br/psdc/cadastro.ConsultaIdentificacaoPessoa.jsf");
					})
					closetx.addEventListener("click", function closeWin(){
					//{
					gwindow.close(0);
					swindow.close(0);
					})
  </script>
     
    
	Em caso de inconsistencia ou falha no carregamento consulte a <a href="http://www.receita.fazenda.gov.br/pessoajuridica/cnpj/cnpjreva/cnpjreva_solicitacao2.asp" target="_blank"> Receita Federal</a>
    <div id="div01"></div>
    <iframe title="Simples Nacional" id="frame" align="right|top"  style="background: #FFFFFF" 	class="iframe" src="http://www8.receita.fazenda.gov.br/SimplesNacional/Aplicacoes/ATBHE/ConsultaOptantes.app/ConsultarOpcao.aspx" onclick="Paste(value);"

</body>
</html>
