import xmind
import pandas as pd
import networkx as nx
import xmd

df = pd.read_excel("层级.xlsx")

G = nx.DiGraph()

for i, row in df.iterrows():
    G.add_edge(row["上级账号"], row["用户账号"])

start_node = "en66"

G.add_node(start_node)

subgraph_nodes = list(nx.dfs_tree(G, start_node).nodes)
subgraph = G.subgraph(subgraph_nodes)

xmind_wb = xmind.load('1.xmind')
sheet = xmind_wb.getPrimarySheet()
sheet.setTitle("层级关系")

root_topic = sheet.getRootTopic()
root_topic.setTitle(start_node)


def add_subtopics(parent_topic, node):
    for child in subgraph.successors(node):
        print(child)
        # child_topic = TopicElement(ownerWorkbook=None)
        child_topic = parent_topic.addSubTopic()
        child_topic.setTitle(child)
        add_subtopics(child_topic, child)


add_subtopics(root_topic, start_node)

xmind.save(xmind_wb, "层级关系.xmind")
xmd.main1("层级关系.xmind")
