import XLSX from 'xlsx'
import { faker } from '@faker-js/faker'
import { cpf } from 'cpf-cnpj-validator'
import path from 'path'
import fs from 'fs'

/** @param {import('@playwright/test').Page} */

const DadosDaPlanilha = (page) => {
  const caminhoDoTemplete = path.resolve('./fixtures/sheet-templete.xlsx')
  const caminhoDoDestino = path.resolve('./fixtures/sheet-nova.xlsx')

  const escreverNaPlanilha = async (dados) => {
    const lerArquivoNoCaminhoSelecionado = fs.readFileSync(caminhoDoTemplete)
    const arquivoExcel = XLSX.read(lerArquivoNoCaminhoSelecionado)

    const nomeDaAba = Object.keys(arquivoExcel.Sheets)[0]
    const abaArquivo = arquivoExcel.Sheets[nomeDaAba]
    const cabecalho = XLSX.utils.sheet_to_json(abaArquivo, { header: 1 })[0]
    const novaAba = XLSX.utils.aoa_to_sheet([cabecalho])

    if (dados.length > 0) {
      XLSX.utils.sheet_add_json(novaAba, dados, {
        skipHeader: true,
        origin: -1,
      })
    }
    arquivoExcel.Sheets[nomeDaAba] = novaAba
    XLSX.writeFile(arquivoExcel, caminhoDoDestino)
    
    return caminhoDoDestino
  }

  const gerarDadosDaTabela = () => {
    return {
      Nome: faker.person.firstName(),
      Sobrenome: faker.person.lastName(),
      CPF: cpf.generate(true),
      Sexo: faker.helpers.arrayElement(['Feminino', 'Masculino']),
    }
  }

  return {
    escreverNaPlanilha,
    gerarDadosDaTabela,
  }
}

const planilha = (page) => DadosDaPlanilha(page)
export default planilha
