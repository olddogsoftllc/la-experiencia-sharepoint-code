// exploreTermStore.js
// Example: Explore the Term Store with Microsoft Graph JavaScript SDK.
// The Graph client is injected via constructor (DI) from the common module.

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
            // Get Term Store
            const termStore = await this.graphClient
                .api('/sites/root/termStore')
                .get();

            console.log('📚 Term Store Info:');
            console.log(`   ID: ${termStore.id}`);
            console.log(`   Default Language: ${termStore.defaultLanguage}`);
            console.log(`   Languages: ${termStore.languageTags?.join(', ') || 'N/A'}`);
            console.log();

            // List groups
            const groups = await this.graphClient
                .api('/sites/root/termStore/groups')
                .get();

            console.log(`📂 Grupos encontrados: ${groups.value.length}\n`);

            for (const group of groups.value) {
                console.log(`   📁 ${group.displayName}`);
                console.log(`      ID: ${group.id}`);
                console.log(`      Description: ${group.description || 'N/A'}`);

                // Get term sets
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

// Example usage: Graph client via common module (client secret, app-only).
async function main() {
    const graphClient = await getGraphClient();
    const explorer = new TermStoreExplorer(graphClient);
    await explorer.explore();
}

module.exports = { TermStoreExplorer };

// If run directly
if (require.main === module) {
    main().catch(console.error);
}
