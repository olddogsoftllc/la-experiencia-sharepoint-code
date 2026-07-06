#!/usr/bin/env python3
# explore_term_store.py
# Example: Explore the Term Store with Microsoft Graph SDK for Python.
# The GraphServiceClient is injected via constructor (DI) from the common module.

import asyncio
from msgraph import GraphServiceClient

from laexperiencia_sharepoint import get_graph_client


class TermStoreExplorer:
    def __init__(self, graph_client: GraphServiceClient):
        self.graph_client = graph_client

    async def explore(self):
        print("=" * 60)
        print("Explorando Term Store")
        print("=" * 60)
        print()

        try:
            # Get Term Store
            term_store = await self.graph_client.sites.by_site_id("root").term_store.get()

            print("📚 Term Store Info:")
            print(f"   ID: {term_store.id}")
            print(f"   Default Language: {term_store.default_language_tag}")
            print(f"   Languages: {', '.join(term_store.language_tags or [])}")
            print()

            # List groups
            groups = await self.graph_client.sites.by_site_id("root").term_store.groups.get()

            print(f"📂 Grupos encontrados: {len(groups.value)}\n")

            for group in groups.value:
                print(f"   📁 {group.display_name}")
                print(f"      ID: {group.id}")
                print(f"      Description: {group.description or 'N/A'}")

                # Get term sets
                sets = await self.graph_client.sites.by_site_id("root") \
                    .term_store.groups.by_group_id(group.id).sets.get()

                for set_item in sets.value:
                    set_name = set_item.localized_names[0].name if set_item.localized_names else set_item.id
                    print(f"         📚 {set_name}")

                print()

        except Exception as e:
            print(f"❌ Error: {e}")


async def main():
    # Graph client via common module (client secret, app-only).
    graph_client = get_graph_client()
    explorer = TermStoreExplorer(graph_client)
    await explorer.explore()


if __name__ == "__main__":
    asyncio.run(main())
