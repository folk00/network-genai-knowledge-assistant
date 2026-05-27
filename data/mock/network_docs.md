# Synthetic Network Knowledge Base

All examples in this file are invented. They are not copied from certification
question banks, customer networks, or proprietary documentation.

---

## DOC-001 - Transit Gateway VPN Routing

Tenant: demo-retail
Site: branch-001
Technology: AWS Transit Gateway, Site-to-Site VPN, BGP

A branch router advertises `172.16.20.0/24` to AWS over two Site-to-Site VPN
tunnels. The VPN is attached to a Transit Gateway. The application VPC CIDR is
`10.42.0.0/16`.

For the VPC workloads to reach the branch LAN, the private VPC route table must
send `172.16.20.0/24` to the Transit Gateway. The Transit Gateway route table
must learn or contain `172.16.20.0/24` pointing to the VPN attachment. With BGP,
route propagation can populate the TGW route table; with static VPN, a static TGW
route is required.

---

## DOC-002 - SD-WAN Tunnel Health Triage

Tenant: demo-retail
Site: branch-077
Technology: Cisco SD-WAN, BFD, OMP

When a Cisco SD-WAN branch has data-plane loss but control connections remain
up, validate BFD sessions, TLOC color, NAT behavior, and policy. If OMP routes
are present but traffic fails, inspect data policy, app-route policy, and return
path symmetry. If BFD is down on only one transport, compare underlay reachability
and NAT traversal on that interface.

---

## DOC-003 - MOP Change Risk

Tenant: demo-retail
Site: dc-01
Technology: Nexus, Port-channel, Firewall, VLAN

A firewall migration MOP should preserve the detected port-channel mode. If the
backup shows static `mode on`, do not generate LACP active unless the migration
design explicitly changes the firewall bundle behavior. Access and trunk members
should not be mixed inside the same logical firewall port-channel unless the
target design documents that mixed behavior.

---

## DOC-004 - Proxy ARP Difference

Tenant: demo-retail
Site: campus-01
Technology: IOS XE, NX-OS

IOS XE commonly has proxy ARP enabled by default on routed interfaces unless
`no ip proxy-arp` is configured. NX-OS behaves differently and may require
explicit `ip proxy-arp` when proxy ARP is intended. Migration tooling should not
blindly copy the absence of `no ip proxy-arp` from IOS XE as if it means the same
thing on Nexus.

---

## DOC-005 - Transit Gateway Multicast Receivers

Tenant: demo-finance
Site: region-001
Technology: AWS Transit Gateway, Multicast, EC2, Security Groups

In AWS Transit Gateway multicast designs, the multicast domain is the control
boundary. Before instances can participate, the relevant VPC attachment and the
subnets that contain participating EC2 network interfaces must be associated
with the multicast domain.

Receivers are registered as multicast group members by network interface ID and
group IP address, unless they join dynamically through IGMP when IGMPv2 support
is enabled. Senders are multicast group sources when static source support is
enabled; sources are also identified by network interface ID and group IP.

Security groups on receiver instances must allow the application traffic that is
delivered to the receivers. For UDP multicast applications, validate the UDP
port and the source instance private IP. IGMP controls membership; it is not the
same thing as allowing the application data plane UDP flow.

Validation checks: confirm multicast domain options, subnet associations, group
members, optional group sources, receiver security group rules, NACLs, and the
application UDP listener on the receiver instances.
