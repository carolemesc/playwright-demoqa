import { test, expect } from '@playwright/test'
import planilha from '../actions/planilha'
import path from 'path'
import fs from 'fs'
import XLSX from 'xlsx'

let page

test.beforeAll(async ({ browser }) => {
  page = await browser.newPage()
})

test.describe('Criar uma planilha', () => {
  test.only('deve ser possível criar uma planilha com novos dados a partir do templete', async () => {
    const Tabela = planilha(page)

    const dados = Tabela.gerarDadosDaTabela()

    const tempFilePath = await Tabela.escreverNaPlanilha([dados])

    const projetoRaiz = path.resolve(__dirname, './fixtures')
    const novoCaminho = path.join(projetoRaiz, 'sheet-nova.xlsx')
  })
})
