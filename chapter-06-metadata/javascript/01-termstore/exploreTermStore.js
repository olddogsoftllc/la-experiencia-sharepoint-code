// exploreTermStore.js
// Ejemplo: Explorar el Term Store con Microsoft Graph JavaScript SDK.
// El cliente de Graph se inyecta por constructor (DI) desde el módulo común.

const { getGraphClient } = require('la-experiencia-sharepoint-code/graphAuth');

class TermStoreExplorer {
    constructor(graphClient) {
        this.graphClient = graphClient;
    }

    async explore() {
        console.log('='.repeat(60));
        console.log('Explorando Term Store');
        console.log('='.repeat(60));
        console.log();

        try {
            // Obtener Term Store
            const termStore = await this.graphClient
                .api('/sites/root/termStore')
                .get();

            console.log('📚 Term Store Info:');
            console.log(`   ID: ${termStore.id}`);
            console.log(`   Default Language: ${termStore.defaultLanguage}`);
            console.log(`   Languages: ${termStore.languageTags?.join(', ') || 'N/A'}`);
            console.log();

            // Listar grupos
            const groups = await this.graphClient
                .api('/sites/root/termStore/groups')
                .get();

            console.log(`📂 Grupos encontrados: ${groups.value.length}\n`);

            for (const group of groups.value) {
                console.log(`   📁 ${group.displayName}`);
                console.log(`      ID: ${group.id}`);
                console.log(`      Descripción: ${group.description || 'N/A'}`);

                // Obtener conjuntos de términos
                const sets = await this.graphClient
                    .api(`/sites/root/termStore/groups/${group.id}/sets`)
                    .get();

                for (const set of sets.value) {
                    const setName = set.localizedNames?.[0]?.name || set.id;
                    console.log(`         📚 ${setName}`);
                }
                console.log();
            }

        } catch (error) {
            console.error('❌ Error:', error.message);
        }
    }
}

// Ejemplo de uso: cliente de Graph vía módulo común (client secret, app-only).
async function main() {
    const graphClient = await getGraphClient();
    const explorer = new TermStoreExplorer(graphClient);
    await explorer.explore();
}

module.exports = { TermStoreExplorer };

// Si se ejecuta directamente
if (require.main === module) {
    main().catch(console.error);
}
