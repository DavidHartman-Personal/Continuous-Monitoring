import io
import re
import cmdb.current_cmdb as current_cmdb

#

SERVER_IP_RE = r"(\S+)\s*\(?.*\)?\s*Ports:\s*((.*)+)"
# SERVER_IP_RE = r"(\S+)\s*Ports:\s*((.*)+)"
STARTS_WITH_AFFECTS_RE = r"\s*Affect.*"
PLUGIN_ID_RE = r"Plugin ID: (\d+)"

server_ip_regex = re.compile(SERVER_IP_RE, re.MULTILINE)
starts_with_affects_regex = re.compile(STARTS_WITH_AFFECTS_RE, re.MULTILINE)
plugin_id_regex = re.compile(PLUGIN_ID_RE, re.MULTILINE)

def split_server_port(in_server_ports):
    """Breaks a single line containing an IP/Server and ports affected into a structure with the name and the ports

    server1 Ports: 0
    server2 Ports: 443
    Server3 Ports: 8099, 443



    """
    server_ports = {}
    server_port_lines = in_server_ports.split('\n')
    line_number = 0
    for line in server_port_lines:
        line_number += 1
        # Check if the first line is "Affect.*" and skip if it is.
        if line_number == 1 and re.match(starts_with_affects_regex, line):
            continue;
        match = re.search(server_ip_regex, line)
        if match:
            #for match in matches:
            match_server_ip = match.group(1)
            server_ip = current_cmdb.get_name_from_alias(match_server_ip.upper())
            server_ports[server_ip] = {'NAME' : str(server_ip),
                                       'ORIGINAL_SERVER' : match.group(1),
                                       'PORTS': []
                                       }
            ports = [item.strip() for item in match.group(2).split(',')]
            ports = [i for i in ports if i]
            for port in ports:
                server_ports[server_ip]['PORTS'].append(port)
        else:
            print("Invalid Server Port definition in Affected Hosts: " + str(in_server_ports))
        # matches = re.finditer(server_ip_regex, line, re.MULTILINE)
        # for match in matches:
        #     server_ip = match.group(1)
        #     server_ports[server_ip] = {'NAME': str(server_ip),
        #                                 'PORTS': []
        #                                 }
        #     ports = [item.strip() for item in match.group(2).split(',')]
        #     ports = [i for i in ports if i]
        #     for port in ports:
        #         server_ports[server_ip]['PORTS'].append(port)
    return server_ports

def get_plugin_id(weakness_source_identifier):
    return_plugin_id = ""
    matches = re.search(plugin_id_regex, weakness_source_identifier)
    if matches:
        return_plugin_id = matches.group(1)
    else:
        return_plugin_id = weakness_source_identifier
    return return_plugin_id

if __name__ == "__main__":
    print("extract-field-info.py")