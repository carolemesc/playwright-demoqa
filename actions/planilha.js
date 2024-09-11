import XLSX from 'xlsx'
import { faker } from '@faker-js/faker'
import { cpf } from 'cpf-cnpj-validator'
import path from 'path'
import fs from 'fs'

/** @param {import('@playwright/test').Page} */

const DadosDaPlanilha = (page) => {
  const caminhoDoTemplete = path.resolve('./fixtures/sheet-nova.xlsx')

  const escreverNaPlanilha = async (dados) => {
    const lerArquivoNoCaminhoSelecionado = fs.readFileSync(caminhoDoTemplete)
    const arquivoExcel = XLSX.read(lerArquivoNoCaminhoSelecionado)

    const nomeDaAba = Object.keys(arquivoExcel.Sheets)[0]
    const abaArquivo = arquivoExcel.Sheets[nomeDaAba]

    XLSX.utils.sheet_add_json(abaArquivo, dados, {
      skipHeader: true,
      origin: -1,
    })
    const path = './fixtures/sheet-nova.xlsx'
    XLSX.writeFile(arquivoExcel, path)
    return path
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
